import enum
import json
import os
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Sequence, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
import xlsxwriter
import openai
import pandas as pd
from sklearn.cluster import AgglomerativeClustering
from tqdm import tqdm
import time

from dotenv import load_dotenv
load_dotenv()  # Automatically loads variables from .env into os.environ



class Approach(str, enum.Enum):
    INDUCTIVE = "inductive"
    DEDUCTIVE = "deductive"
    HYBRID = "hybrid"


@dataclass
class Code:
    label: str
    quote: str
    interview_id: str


@dataclass
class Theme:
    name: str
    definition: str
    codes: List[Code] = field(default_factory=list)


class ThematicAnalyzer:
    def __init__(
        self,
        approach: Approach = Approach.INDUCTIVE,
        model: str = "gpt-4o-mini",  # CHEAP default
        system_prompt: Optional[str] = None,
        temperature: float = 0.2,
        max_tokens: int = 1024,
        save_cache: Optional[str] = None,
    ) -> None:
        self.approach = approach
        self.model = model
        self.temperature = temperature
        self.max_tokens = max_tokens
        self._system_prompt = system_prompt or (
            "You are an expert qualitative researcher performing Braun & Clarke thematic analysis."
        )

        self.interviews: Dict[str, str] = {}
        self.codebook: List[str] = []
        self.research_questions: List[str] = []
        self.themes: List[Theme] = []
        self.custom_rows: Dict[str, str] = {}  # row_label -> prompt

        self._cache_path = Path(save_cache) if save_cache else None
        if self._cache_path:
            self._cache_path.parent.mkdir(parents=True, exist_ok=True)

        openai.api_key = os.getenv("OPENAI_API_KEY")

    # ------------------------------------------------------------------
    # Interview ingestion helpers
    # ------------------------------------------------------------------

    def add_interview(self, interview_id: str, transcript: str) -> None:
        self.interviews[interview_id] = transcript

    def load_transcripts_from_dir(self, folder, exts=(".txt", ".md")) -> None:
        """
        Recursively read transcripts.  
        • filename stem (e.g. P01) becomes interview-ID  
        • if any folder name contains “kenya” (case-insensitive) → region = Kenya  
        otherwise default = Zambia  
        • final stored ID looks like  P01|Kenya   or  P07|Zambia
        """
        folder = Path(folder)
        for p in folder.rglob("*"):
            if p.suffix.lower() in exts:
                iid = p.stem
                region = "Kenya" if any("kenya" in part.lower() for part in p.parts) else "Zambia"
                self.add_interview(f"{iid}|{region}", p.read_text(encoding="utf-8"))

    # ------------------------------------------------------------------
    # Config setters
    # ------------------------------------------------------------------

    def set_codebook(self, codes: Sequence[str]):
        self.codebook = list(codes)

    def set_research_questions(self, questions: Sequence[str]):
        self.research_questions = list(questions)

    def add_custom_row(self, label: str, prompt: str):
        """Register a researcher‑defined analytic row (answered per interview)."""
        self.custom_rows[label] = prompt

    # ------------------------------------------------------------------
    # Full pipeline driver
    # ------------------------------------------------------------------

    def run(self):
        if not self.interviews:
            raise ValueError("No interviews loaded. Use add_interview() or load_transcripts_from_dir().")

        segmented = {iid: self._segment_transcript(t) for iid, t in self.interviews.items()}
        codes_inductive = (
            self._generate_initial_codes(segmented) if self.approach != Approach.DEDUCTIVE else []
        )
        codes_all = codes_inductive + [Code(cb, "", "*") for cb in self.codebook]
        initial_themes = self._cluster_codes(codes_all)
        self.themes = self._define_and_name_themes(self._review_themes(initial_themes))
        self._assign_rqs()       

    # ------------------------------------------------------------------
    # Excel export with resume capability
    # ------------------------------------------------------------------

    def to_excel(
        self,
        path: str | Path,
        subcolumns: Sequence[str] | None = None,
        max_quotes_per_cell: int = 3,
        resume: bool = True,
    ) -> None:
        path = Path(path)
        subcolumns = list(subcolumns or ["introduction", "codes", "supporting_quotes"])

        # Build fresh dataframe
        cols_themes = pd.MultiIndex.from_product([self.interviews.keys(), subcolumns])
        cols_themes = pd.MultiIndex.from_tuples([("Theme Definition", "")] + list(cols_themes))
        rows = []
        for t in self.themes:
            rqs = t.rqs or ["Unassigned"]
            for rq in rqs:
                rows.append( (rq, t.name) )

        index = pd.MultiIndex.from_tuples(rows, names=["Research Question","Theme"])
        df_themes = pd.DataFrame(index=index, columns=cols_themes)
        df_themes.sort_index(inplace=True)

        # Collect top quotes and build cell content (codes and supporting quotes)
        intro_jobs = []
        for theme in tqdm(self.themes, desc="Processing themes"):
            for rq in theme.rqs or ["Unassigned"]:
                df_themes.loc[(rq, theme.name), ("Theme Definition", "")] = theme.definition
                for iid in tqdm(self.interviews, leave=False, desc=f"Processing {theme.name}"):
                    codes_here = [c for c in theme.codes if c.interview_id == iid and c.quote]
                    if not codes_here:
                        continue
                    top_codes = self._select_representative_quotes(codes_here, theme.definition, max_quotes_per_cell)
                    for code in top_codes:
                        if "codes" in subcolumns:
                            df_themes.loc[(rq, theme.name), (iid, "codes")] = _acc_codes(
                                df_themes.loc[(rq, theme.name), (iid, "codes")], code.label
                            )
                        if "supporting_quotes" in subcolumns:
                            df_themes.loc[(rq, theme.name), (iid, "supporting_quotes")] = _acc(
                                df_themes.loc[(rq, theme.name), (iid, "supporting_quotes")], code.quote, sep="\n"
                            )

                if "introduction" in subcolumns:
                    intro_jobs.append((theme.name, theme.definition, iid, [c.quote for c in top_codes]))

        # Parallel intro generation per theme × interview
        from concurrent.futures import ThreadPoolExecutor

        intro_results = {}
        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = [
                executor.submit(self._generate_intro_batch, tname, tdef, quotes)
                for tname, tdef, _, quotes in intro_jobs
            ]
            for (tname, _, iid, _), future in tqdm(zip(intro_jobs, futures), total=len(intro_jobs), desc="Generating intros"):
                try:
                    intro_results[(tname, iid)] = future.result()
                except Exception as e:
                    print(f"[!] Failed to generate intro for ({tname}, {iid}): {e}")
                    intro_results[(tname, iid)] = ""

        # Assign intros to dataframe
        for (tname, iid), intro in intro_results.items():
            df_themes.loc[tname, (iid, "introduction")] = intro

        # Compute custom analytic rows
        df_custom = None
        if self.custom_rows:
            rows = {}
            for row_lbl, prompt_tpl in tqdm(self.custom_rows.items(), desc="Generating custom rows"):
                for iid, txt in tqdm(self.interviews.items(), desc=f"Custom row: {row_lbl}", leave=False):
                    answer = self._chat([
                        {"role": "user", "content": prompt_tpl.replace("{transcript}", txt)}
                    ])
                    rows.setdefault(row_lbl, {})[iid] = answer.strip()
            df_custom = pd.DataFrame(rows).T
            # Add empty multilevel columns for consistency
            df_custom.columns = pd.MultiIndex.from_tuples([(iid, "answer") for iid in df_custom.columns])

        # Resume merge
        if resume and path.exists():
            with pd.ExcelFile(path) as xls:
                old = pd.read_excel(xls, sheet_name="Themes × Interviews", header=[0, 1], index_col=0)
                df_themes = _merge_dfs(old, df_themes)
                if df_custom is not None and "Custom Analyses" in xls.sheet_names:
                    old_custom = pd.read_excel(xls, sheet_name="Custom Analyses", header=[0, 1], index_col=0)
                    df_custom = _merge_dfs(old_custom, df_custom)

        # Write to workbook
        with pd.ExcelWriter(path, engine="xlsxwriter") as xl:
            wrote_themes = False

            # Only write theme sheet if relevant
            if self.themes and subcolumns:
                df_themes.to_excel(xl, sheet_name="Themes × Interviews")
                wrote_themes = True

            if df_custom is not None:
                df_custom.to_excel(xl, sheet_name="Custom Analyses")
                stats = {}
                for row in df_custom.index:
                    answers = [
                        f"{col[0]}: {df_custom.at[row, col]}"
                        for col in df_custom.columns
                        if pd.notna(df_custom.at[row, col])
                    ]
                    if not answers:
                        continue
                    prompt = (
                        f"You are analyzing answers to the question:\n\n"
                        f"\"{row}\"\n\n"
                        "Write a concise 1–3 sentence summary that quantifies patterns and highlights key themes, while explaining the rationale and mention who said what when relevant to explain why they said it.\n"
                        "Use short and clear language, e.g. 'Most respondents... (8/12)', 'Others (2/12)', or 'Key challenge: XYZ'.\n\n"
                        "Responses:\n" + "\n".join(answers)
                    )
                    response = self._chat([{"role": "user", "content": prompt}])
                    stats[row] = {"Analytical Summary": response.strip()}
                df_stats = pd.DataFrame(stats).T
                df_stats.to_excel(xl, sheet_name="Summary Stats")

                for region in ("Kenya", "Zambia"):
                    cols = [c for c in df_custom.columns if f"|{region}" in c[0]]
                    if not cols:
                        continue
                    sub = df_custom[cols]
                    region_stats = {}
                    for row in sub.index:
                        answers = [
                            f"{col[0]}: {sub.at[row, col]}"
                            for col in sub.columns
                            if pd.notna(sub.at[row, col])
                        ]
                        if not answers:
                            continue
                        prompt_region = (
                            f"You are analyzing short interview responses from {region} only.\n"
                            f"You are analyzing answers to the question:\n\n"
                            f"\"{row}\"\n\n"
                            "Write a concise 1–3 sentence summary that quantifies patterns and highlights key themes, while explaining the rationale and mention who said what when relevant to explain why they said it.\n"
                            "Use short and clear language, e.g. 'Most respondents... (10/12)', 'Others (2/12)', or 'Key challenge: XYZ'.\n\n"
                            f"Responses from {region}:\n" + "\n".join(answers)
                        )
                        response = self._chat([{"role": "user", "content": prompt_region}])
                        region_stats[row] = {f"{region} Analytical Summary": response.strip()}
                    pd.DataFrame(region_stats).T.to_excel(xl, sheet_name=f"Stats – {region}")

            if wrote_themes:
                # Now it's safe to access df_themes
                ws = xl.sheets["Themes × Interviews"]
                colour1, colour2 = "#F0F5FF", "#FFFFFF"

                if isinstance(df_themes.index, pd.MultiIndex):
                    rq_levels = df_themes.index.get_level_values(0).unique()
                else:
                    rq_levels = df_themes.index.unique()

                for i, rq in enumerate(rq_levels):
                    matches = [j for j, idx in enumerate(df_themes.index) if isinstance(idx, tuple) and idx[0] == rq]
                    if not matches:
                        continue
                    start = matches[0] + 1
                    end = matches[-1] + 1
                    fmt = xl.book.add_format({"bg_color": colour1 if i % 2 == 0 else colour2})
                    for rownum in range(start, end + 1):
                        ws.set_row(rownum, cell_format=fmt)

                deep_insights = {}
                for theme in tqdm(self.themes, desc="Generating thematic insights"):
                    kenya_quotes = []
                    zambia_quotes = []

                    for iid in self.interviews:
                        for rq in theme.rqs or ["Unassigned"]:
                            cell = df_themes.loc[(rq, theme.name), (iid, "supporting_quotes")]
                            if pd.notna(cell) and isinstance(cell, str):
                                named_quotes = f"{iid}: {cell.strip()}"
                                if "Kenya" in iid:
                                    kenya_quotes.append(named_quotes)
                                elif "Zambia" in iid:
                                    zambia_quotes.append(named_quotes)

                    prompt_theme = (
                        f"You are analyzing qualitative interview data for the theme '{theme.name}'.\n"
                        f"Definition: {theme.definition}\n\n"
                        "Write a detailed analytical summary comparing how this theme is reflected in Zambia vs. Kenya. Make it interesting, focus on contrasts and surprising patterns. Don't focus on boring or vague quotes.\n"
                        "Refer to interviewees by name when highlighting insights. Use counts where possible, e.g. (5/12) said X.\n\n"
                        f"Quotes from Zambia:\n{json.dumps(zambia_quotes)}\n\n"
                        f"Quotes from Kenya:\n{json.dumps(kenya_quotes)}"
                    )
                    insight = self._chat([{"role": "user", "content": prompt_theme}])
                    deep_insights[theme.name] = insight.strip()

                df_insights = pd.DataFrame.from_dict(deep_insights, orient="index", columns=["Cross-Country Summary"])
                df_insights.to_excel(xl, sheet_name="Thematic Insights")

    # ------------------------------------------------------------------
    # Internal helpers (segment, GPT chat, coding, clustering etc.)
    # ------------------------------------------------------------------

    def _segment_transcript(self, transcript: str, max_chars: int = 600) -> List[str]:
        sents = re.split(r"(?<=[.!?])\s+", transcript)
        chunk, out = "", []
        for s in sents:
            if len(chunk) + len(s) < max_chars:
                chunk += " " + s
            else:
                out.append(chunk.strip())
                chunk = s
        if chunk:
            out.append(chunk.strip())
        return out

    def _chat(self, messages: List[dict], max_retries: int = 100, backoff_factor: float = 1.5) -> str:
        retries = 0
        delay = 1  # Initial delay in seconds
        while retries < max_retries:
            try:
                response = openai.chat.completions.create(
                    model=self.model,
                    messages=[{"role": "system", "content": self._system_prompt}] + messages,
                    temperature=self.temperature,
                    max_tokens=self.max_tokens,
                )
                text = response.choices[0].message.content.strip()
                if self._cache_path:
                    with self._cache_path.open("a", encoding="utf-8") as fh:
                        fh.write(json.dumps({"messages": messages, "response": text}) + "\n")
                return text
            except openai.error.RateLimitError:
                print(f"Rate limit exceeded. Retrying in {delay} seconds...")
                time.sleep(delay)
                delay *= backoff_factor
                retries += 1
            except Exception as e:
                print(f"An error occurred: {e}")
                raise
        raise Exception("Maximum retries exceeded due to rate limiting.")
    
    def _embedding(self, text: str, max_retries: int = 100, backoff_factor: float = 1.5) -> List[float]:
        retries = 0
        delay = 1
        while retries < max_retries:
            try:
                resp = openai.embeddings.create(input=[text], model="text-embedding-3-small")
                return resp.data[0].embedding
            except openai.error.RateLimitError:
                print(f"[Embedding] Rate limit hit. Retrying in {delay} seconds...")
                time.sleep(delay)
                delay *= backoff_factor
                retries += 1
            except Exception as e:
                print(f"[Embedding] Error: {e}")
                raise
        raise Exception("Embedding retries exceeded.")

    def _generate_initial_codes(self, segmented: Dict[str, List[str]]) -> List[Code]:
        rq_context = "\n\nResearch Questions:\n" + "\n".join(self.research_questions) if self.research_questions else ""

        prompts = []
        ids_and_chunks = []
        for iid, chunks in segmented.items():
            for chunk in chunks:
                prompt_text = f"Identify concise semantic codes (max 3) that relate to the excerpt. Return JSON list of strings only.\n{rq_context}\n\n{chunk}"
                prompts.append({"role": "user", "content": prompt_text})
                ids_and_chunks.append((iid, chunk))

        codes: List[Code] = []

        def call_gpt(message: dict) -> List[str]:
            retries = 0
            delay = 1.0
            max_retries = 100
            backoff = 1.5

            while retries < max_retries:
                try:
                    response = openai.chat.completions.create(
                        model=self.model,
                        messages=[{"role": "system", "content": self._system_prompt}, message],
                        temperature=self.temperature,
                        max_tokens=self.max_tokens,
                    )
                    content = response.choices[0].message.content.strip()

                    if not content:
                        print("[!] Empty response from GPT")
                        return []

                    if content.startswith("```") and content.endswith("```"):
                        content = re.sub(r"^```(?:json)?\n|\n```$", "", content.strip())

                    try:
                        return json.loads(content)
                    except json.JSONDecodeError:
                        print(f"[!] Failed to decode JSON:\n{content}")
                        return []

                except openai.error.RateLimitError:
                    print(f"[!] Rate limit error. Retrying in {delay:.1f}s...")
                    time.sleep(delay)
                    delay *= backoff
                    retries += 1

                except KeyboardInterrupt:
                    print("Aborted by user. Shutting down...")
                    return []

                except Exception as e:
                    print(f"[!] GPT error: {e}")
                    return []

            raise Exception("[!] call_gpt: Exceeded maximum retries")
            

        with ThreadPoolExecutor(max_workers=20) as executor:
            futures = [
                executor.submit(call_gpt, prompt)
                for prompt in prompts
            ]
            for future, (iid, chunk) in tqdm(zip(as_completed(futures), ids_and_chunks), total=len(futures), desc="Initial coding"):
                labels = future.result()
                for lab in labels:
                    codes.append(Code(label=lab, quote=chunk, interview_id=iid))

        return codes

    def _cluster_codes(self, codes: List[Code]) -> List[Theme]:
        if not codes:
            return []

        labels = list({c.label for c in codes})

        def batch_embed(texts: List[str]) -> List[List[float]]:
            retries = 0
            delay = 1.0
            max_retries = 100
            backoff = 1.5

            while retries < max_retries:
                try:
                    response = openai.embeddings.create(input=texts, model="text-embedding-3-small")
                    return [r.embedding for r in response.data]
                except openai.error.RateLimitError:
                    print(f"[!] Embedding rate limit hit. Retrying in {delay:.1f}s...")
                    time.sleep(delay)
                    delay *= backoff
                    retries += 1
                except Exception as e:
                    print(f"[!] Embedding error: {e}")
                    return [[0.0] * 1536 for _ in texts]  # Fallback: dummy vector
            raise Exception("[!] batch_embed: Exceeded maximum retries")

        # Split into batches for efficiency (OpenAI allows batch inputs!)
        batch_size = 100
        batches = [labels[i:i+batch_size] for i in range(0, len(labels), batch_size)]
        
        embeddings: List[List[float]] = []
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = [executor.submit(batch_embed, b) for b in batches]
            for f in tqdm(futures, desc="Embedding codes", total=len(futures)):
                embeddings.extend(f.result())

        from sklearn.cluster import AgglomerativeClustering

        n_clusters = max(2, 12)
        clusterer = AgglomerativeClustering(n_clusters=n_clusters)
        assignments = clusterer.fit_predict(embeddings)

        cluster_dict: Dict[int, Theme] = {}
        for lab, cl in zip(labels, assignments):
            if cl not in cluster_dict:
                cluster_dict[cl] = Theme(name=f"Theme {cl+1}", definition="")
            cluster_dict[cl].codes.extend([c for c in codes if c.label == lab])

        return list(cluster_dict.values())

    def _review_themes(self, themes: List[Theme]) -> List[Theme]:
        summary = "\n".join([f"{t.name}: {[c.label for c in t.codes]}" for t in themes])
        prompt = (
            "Given these preliminary themes and their codes, suggest merges, splits, or removals. "
            "Return revised theme groups as JSON mapping theme_name -> list[codes].\n\n" + summary
        )
        revised_json = self._chat([{"role": "user", "content": prompt}])
        try:
            mapping = json.loads(revised_json)
        except json.JSONDecodeError:
            return themes

        name_to_codes = defaultdict(list)
        for old in themes:
            for cd in old.codes:
                name_to_codes[cd.label].append(cd)

        out: List[Theme] = []
        for tname, clist in mapping.items():
            theme = Theme(name=tname, definition="")
            for label in clist:
                theme.codes.extend(name_to_codes.get(label, []))
            out.append(theme)
        return out
    
    def _assign_rqs(self) -> None:
        """Populate theme.rqs with ['RQ1', 'RQ3', …] via ONE GPT call per theme."""
        if not self.research_questions:
            for t in self.themes:
                t.rqs = []
            return

        rq_list = "\n".join(self.research_questions)
        for t in self.themes:
            prompt = (
                "Below are the research questions and a theme definition.\n"
                "Return a JSON list of RQ identifiers (e.g., ['RQ1','RQ3']) this theme addresses.\n\n"
                f"Research Questions:\n{rq_list}\n\nTheme:\n{t.definition}"
            )
            reply = self._chat([{"role":"user","content":prompt}])
            try:
                t.rqs = json.loads(reply)
            except Exception:
                t.rqs = []

    def _define_and_name_themes(self, themes: List[Theme]) -> List[Theme]:
        def define(theme: Theme) -> Theme:
            quotes = [c.quote for c in theme.codes if c.quote][:5]
            prompt = (
                "Define the overarching concept that unites these codes and assign a name. "
                "Return JSON with 'name' and a <= 40 word'definition'.\n\nCodes: "
                + str([c.label for c in theme.codes]) + "\nQuotes: " + str(quotes)
            )
            reply = self._chat([{"role": "user", "content": prompt}])

            # Clean up markdown fences
            if reply.startswith("```") and reply.endswith("```"):
                reply = re.sub(r"^```(?:json)?\n|\n```$", "", reply.strip())

            try:
                obj = json.loads(reply)
                theme.name = obj.get("name", theme.name)
                theme.definition = obj.get("definition", "")
            except json.JSONDecodeError:
                print(f"[!] Failed to define theme:\n{reply}")
            return theme

        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = [executor.submit(define, theme) for theme in themes]
            themes = [f.result() for f in tqdm(futures, desc="Naming themes", total=len(themes))]

        return themes

    def refine_codes(self, method: str = "gpt") -> None:
        if method == "gpt":
            all_labels = list({c.label for t in self.themes for c in t.codes})
            prompt = (
                "Group the following semantic code labels into 10–20 meaningful categories. "
                "Return a JSON object mapping original_label -> normalized_label.\n\nLabels:\n"
                + json.dumps(all_labels)
            )
            reply = self._chat([{"role": "user", "content": prompt}])

            if not reply:
                print("[!] Empty response from GPT.")
                return

            # Clean ```json\n...\n``` blocks and explanations
            if "```" in reply:
                match = re.search(r"```(?:json)?\s*(.*?)\s*```", reply, re.DOTALL)
                if match:
                    reply = match.group(1).strip()

            if not reply.strip().startswith("{"):
                match = re.search(r"\{.*\}", reply, re.DOTALL)
                if match:
                    reply = match.group(0).strip()

            try:
                mapping = json.loads(reply)
                for theme in self.themes:
                    for code in theme.codes:
                        if code.label in mapping:
                            code.label = mapping[code.label]
            except json.JSONDecodeError:
                print("[!] Failed to parse GPT normalization response:\n", reply)

        elif method == "cosine":
            from sklearn.metrics.pairwise import cosine_similarity
            import numpy as np

            all_labels = list({c.label for t in self.themes for c in t.codes})

            # Parallel embedding generation
            def embed(label: str):
                try:
                    return self._embedding(label)
                except Exception as e:
                    print(f"[!] Embedding failed for '{label}': {e}")
                    return [0.0] * 1536  # fallback vector

            with ThreadPoolExecutor(max_workers=10) as executor:
                embeddings = list(executor.map(embed, all_labels))

            sim_matrix = cosine_similarity(embeddings)
            label_map = {}

            for i, label in tqdm(enumerate(all_labels), total=len(all_labels), desc="Cosine similarity grouping"):
                if label in label_map:
                    continue
                for j in range(i + 1, len(all_labels)):
                    if sim_matrix[i][j] > 0.9:
                        label_map[all_labels[j]] = label

            for theme in self.themes:
                for code in theme.codes:
                    if code.label in label_map:
                        code.label = label_map[code.label]

    def _select_representative_quotes(self, codes: List[Code], theme_def: str, max_n: int = 3) -> List[Code]:
        """Use GPT to select the most illustrative quotes for a theme."""
        if len(codes) <= max_n:
            return codes

        prompt = (
            "You are assisting in qualitative research. Given the following theme definition and quotes, "
            "select the top quotes (max {}) that best illustrate the theme. Return a JSON list of selected quotes verbatim.\n\n"
            "Theme Definition:\n{}\n\nQuotes:\n{}"
        ).format(max_n, theme_def, json.dumps([c.quote for c in codes]))

        reply = self._chat([{"role": "user", "content": prompt}])
        try:
            selected_quotes = json.loads(reply)
            return [c for c in codes if c.quote in selected_quotes][:max_n]
        except json.JSONDecodeError:
            return codes[:max_n]  # fallback

    def _generate_intro_batch(self, theme_name: str, theme_def: str, quotes: List[str]) -> str:
        if not quotes:
            return ""
        prompt = (
            f"You are writing a short introduction for a thematic analysis report.\n"
            f"Theme: {theme_name}\nDefinition: {theme_def}\n\n"
            f"Given the following quotes related to this theme, give one sentence (max 25 words) summarizing what they collectively illustrate.\n\n"
            f"Quotes:\n{json.dumps(quotes)}"
        )
        reply = self._chat([{"role": "user", "content": prompt}])
        if reply.startswith("```") and reply.endswith("```"):
            reply = re.sub(r"^```(?:json)?\n|\n```$", "", reply.strip())
        return reply.strip()
    


def _acc(existing: str | float | None, new: str, sep: str = "; ") -> str:
    if existing and not pd.isna(existing):
        return str(existing) + sep + new
    return new


def _acc_codes(existing: str | float | None, new: str, sep="; ") -> str:
    """Merge code labels into a set so duplicates never appear."""
    s = set()
    if isinstance(existing, pd.Series):
        existing = existing.iloc[0] if not existing.empty else None
    if isinstance(existing, str):
        s.update(existing.split(sep))
    elif pd.notna(existing):
        s.add(str(existing))
    s.add(new)
    return sep.join(sorted(filter(None, s)))

def _safe_json_parse(response: str, fallback=None) -> any:
    """
    Safely parse a JSON string. Returns fallback on failure.
    Also cleans ```json blocks.
    """
    if not response or not isinstance(response, str):
        return fallback
    response = response.strip()
    if response.startswith("```") and response.endswith("```"):
        response = re.sub(r"^```(?:json)?\n|\n```$", "", response.strip())
    try:
        return json.loads(response)
    except json.JSONDecodeError:
        print("[!] Failed to parse JSON:\n", response)
        return fallback

# ------------------------------------------------------------------
# Merge helper – keeps existing data, appends new rows / columns,
# and updates cells that were previously empty
# ------------------------------------------------------------------
def _merge_dfs(df_old: pd.DataFrame, df_new: pd.DataFrame) -> pd.DataFrame:
    """
    Combine two multilevel-column dataframes (same shape as our Excel sheets).

    • Rows/themes/interviews found only in df_new are appended.
    • Columns (new interviews or new sub-columns) found only in df_new are appended.
    • For cells that exist in both, we keep the old value unless the old cell is
    empty/NaN, in which case we take the new value.  If *both* contain text we
    concatenate them with a line-break so nothing is lost.
    """
    merged = df_old.copy()

    # 1. Add any completely new columns
    for col in df_new.columns:
        if col not in merged.columns:
            merged[col] = pd.NA

    # 2. Add any completely new rows
    new_rows = df_new.index.difference(merged.index)
    if len(new_rows):
        merged = pd.concat([merged, df_new.loc[new_rows]], axis=0)

    # 3. Update/concatenate overlapping cells
    common_rows = df_new.index.intersection(merged.index).tolist() + df_new.index.difference(merged.index).tolist()
    for idx in df_new.index:
        for col in df_new.columns:
            new_val = df_new.loc[idx, col]
            if pd.isna(new_val) or new_val == "" or new_val is None:
                continue

            # Ensure row exists — works for MultiIndex too
            if idx not in merged.index:
                new_row = pd.DataFrame(
                    [pd.NA], index=pd.MultiIndex.from_tuples([idx], names=merged.index.names),
                    columns=pd.MultiIndex.from_tuples([col], names=merged.columns.names)
                )
                merged = pd.concat([merged, new_row], axis=0)

            # Ensure column exists
            if col not in merged.columns:
                merged[col] = pd.NA

            # Safe value access and update
            old_val = merged.loc[idx, col]
            if pd.isna(old_val) or old_val == "" or old_val is None:
                merged.loc[idx, col] = new_val
            elif str(new_val) not in str(old_val):
                merged.loc[idx, col] = f"{old_val}\n{new_val}"

    # Re-sort rows so existing order is preserved and new ones come last
    merged = merged.loc[df_old.index.tolist() + new_rows.tolist()]
    return merged



# Optional CLI usage
if __name__ == "__main__":
    import argparse, textwrap

    ep = "Minimal command-line wrapper around ThematicAnalyzer."
    parser = argparse.ArgumentParser(ep, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("transcripts", nargs="+", help="paths to interview .txt files")
    parser.add_argument("--approach", choices=[e.value for e in Approach], default="inductive")
    parser.add_argument("--output", default="analysis.xlsx")
    args = parser.parse_args()

    ta = ThematicAnalyzer(approach=Approach(args.approach))
    for p in args.transcripts:
        iid = Path(p).stem
        ta.add_interview(iid, Path(p).read_text())
    ta.run()
    ta.refine_codes(method="gpt")
    ta.to_excel(args.output, max_quotes_per_cell=1)
    print(f"Saved results to {args.output}")
