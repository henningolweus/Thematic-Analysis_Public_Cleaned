from thematic_analyzer import ThematicAnalyzer, Approach
import sys
from pathlib import Path

# 1. Set up the analyzer
ta = ThematicAnalyzer(approach=Approach.INDUCTIVE, model="gpt-4o-mini")

# 2. Load all .txt files from the 'transcripts' folder
ta.load_transcripts_from_dir("transcripts")

# 3. Add custom questions

ta.add_custom_row(

    "Ticket Sizes & Stage Coverage",
    """Does this transcript mention **ticket sizes / cheque ranges** and the **investment stages** (e.g., pre-seed, seed, Series A/B/C) for
  (a) the interviewee’s **own fund**, and/or  
  (b) the **wider venture-investment market (PE/VC/grants or other)** in Zambia or Kenya?

If **yes**, summarise in two labelled parts:

**Own fund:** <brief summary of ticket range & stages>.  
**Market:** <brief summary of ranges & stages mentioned for Zambia/Kenya as a whole>.

For each part you include, add one or more **supporting quotes** that clearly state the figures or stage labels (verbatim from the transcript).

If **no ticket-size or stage information** is provided anywhere in the transcript, write:  
"Not mentioned in this interview."

Transcript:
{transcript}"""
)


ta.add_custom_row(
    "Capital Stack Gaps and Market dynamics",
    """Does this transcript mention anything about the **Depth of the Financial Markets** (e.g., collaboration between different investors/accellerators and other actors, investor appetite, etc.) or **constraints in exit opportunities** (e.g., weak IPO markets, few trade sales, long holding periods) in Kenya or Zambia?

If yes, summarize how the interviewee describes these limitations— investor appetite. Include one to many unedited**supporting quotes** that clearly illustrates the capital market dynamics or exit challenge.

If no such gaps or constraints are discussed, write: "Not mentioned in this interview."

Transcript:
{transcript}"""
)

# 4. Export just the custom questions
ta.to_excel("custom_questions_and_answers.xlsx", subcolumns=[], resume=True)