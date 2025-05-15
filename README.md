✅ README.md
markdown
Kopier
Rediger
# 🎯 Thematic-Analyzer

**Automated Braun & Clarke thematic analysis for batches of interview transcripts – powered by OpenAI.**  
You drop your `.txt` or `.md` transcripts into a folder, run one command, and get a multi-sheet Excel workbook with:

| Sheet                  | What’s inside |
|------------------------|---------------|
| **Themes × Interviews** | Theme definitions, per-interview codes, and the representative quotes you told the tool to keep (default: ≤3 per cell). |
| **Custom Analyses**    | One row per “custom question” you register (e.g. *Ticket Sizes & Stage Coverage*). Each interview’s answer appears in its own column. |
| **Summary Stats**      | A 1–3 sentence GPT summary that quantifies patterns across **all** interviews for every custom question. |
| **Stats – Kenya / Zambia** | Same as above but split by region. |
| **Thematic Insights**  | GPT narration that compares how each theme plays out in Kenya vs. Zambia, with named quotes. |

All long-running OpenAI calls come with exponential-backoff retry logic, so you rarely crash even on free-tier rate limits.

---

## ⚡ Quick start

```bash
git clone https://github.com/your-org/thematic-analyzer.git
cd thematic-analyzer
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
export OPENAI_API_KEY="sk-..."  # or put it in a .env file
Put transcripts in transcripts/.
If a parent folder contains the word kenya (case-insensitive), the file is tagged “|Kenya”; otherwise “|Zambia”.

text
Kopier
Rediger
transcripts/
├── kenya/
│   ├── 04_03_25 - VC firm.txt
│   └── 20_02_25 - PE fund.txt
└── zambia/
    ├── 10_02_25 - Accelerator Hub.txt
    └── 11_02_25 - VC Debt Fund.txt
▶️ Run the full pipeline
bash
Kopier
Rediger
python run.py
# → analysis.xlsx
▶️ (Optional) Only ask custom questions
bash
Kopier
Rediger
python run_custom_questions.py
# → custom_questions_and_answers.xlsx
🛠 What the scripts do
File	Purpose	Key switches
thematic_analyzer.py	Library class that orchestrates everything (segmentation → inductive coding → clustering → GPT theme naming → Excel export).	model, temperature, max_tokens, max_quotes_per_cell, etc.
run.py	End-to-end thematic analysis. Adds demo questions, runs ta.run(), clusters codes, and writes Excel.	Edit or add ta.add_custom_row(...) or change clustering method.
run_custom_questions.py	Skips theme generation and answers only your custom rows.	Use when you just want a structured Q&A sheet.

🚀 Sample output
<details><summary>Custom Analyses → “Ticket Sizes & Stage Coverage” (excerpt)</summary>
Interview	Answer
04_03_25 – VC firm|Kenya	Own fund: $3–20M (sweet spot > $5M), usually Series B.
Market: Kenya still needs patient pre-seed/Series A capital.
Quote: “Our ticket sizes are between $3 M and 20 M … we’d be happy to do >$5 M.”
10_02_25 – Accelerator Hub|Zambia	Own fund: Not mentioned.
Market: Zambia’s VC ecosystem is tiny; ticket sizes rare >$500k.
Quote: “There’s little to no early-stage capital. A typical Zambian enterprise would be looking for around $250k.”

</details> <details><summary>Thematic Insights sheet (one row per theme)</summary>
Theme: Financial Dynamics and Constraints

🇿🇲 Zambia (7/13) stress the absence of local exits and default to debt; 🇰🇪 Kenya (10/18) more often juggle equity + mezzanine but fear FX swings.
Kukula (ZM) calls local banks “collateral-obsessed”, while Simon (KE) lauds SAFEs but still “gets crushed” by shilling weakness.

</details>
🧩 How it works (under the hood)
Segmentation – each transcript is split into ~600-character chunks.

Inductive coding – GPT returns ≤3 semantic codes (JSON) per chunk.

Clustering – codes are embedded (OpenAI text-embedding-3-small) and clustered.

Theme naming – GPT summarizes cluster labels and quotes to define a theme.

RQ tagging (optional) – maps themes to your Research Questions.

Representative quotes – selects top n quotes per theme and interview.

Custom rows – analyst-defined prompts answered per transcript.

Excel writer – merges data into Excel and color-codes by Research Question.

Cross-country narration – GPT writes “Zambia vs Kenya” mini-essays per theme.

All GPT/embedding calls are wrapped with retry logic:

delay = 1
for attempt in range(100):
    try:
        return openai.chat.completions.create(...)
    except openai.error.RateLimitError:
        time.sleep(delay)
        delay *= 1.5
⏱ Runtime & rate limits
Model	Free-tier limit	Worst-case usage (30 transcripts)	Safe?
gpt-4o-mini	~10 req/min, 500/day	~1 req/min	✔
text-embedding-3-small	~250 req/min	batched (100/req)	✔

If limits are hit, retry logic kicks in and continues automatically.

📑 FAQ
Can I switch to a faster model?
Yes – pass model="gpt-3.5-turbo" when creating the ThematicAnalyzer. It’s cheaper, but theme naming quality drops.

How do I change number of quotes per theme/interview?
Set max_quotes_per_cell=1 in ta.to_excel(...).

Where are raw GPT conversations saved?
Use save_cache="logs/gpt_cache.ndjson" when running.

What if a transcript mentions neither Kenya nor Zambia?
Defaults to “Zambia” (see load_transcripts_from_dir).

✨ Contributing
Pull requests welcome – especially for:

new clustering back-ends

better resume logic

additional output formats (e.g. PowerPoint)

📜 License
MIT – do whatever you want, but attribution is appreciated.

Happy coding & analysing! 🕵️‍♀️