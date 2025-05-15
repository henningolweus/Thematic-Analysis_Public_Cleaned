

from thematic_analyzer import ThematicAnalyzer, Approach

# 1. Initialize the analyzer with your chosen approach
ta = ThematicAnalyzer(approach=Approach.INDUCTIVE, model="gpt-4o-mini")

# 2. Load all .txt files from the 'transcripts' folder
ta.load_transcripts_from_dir("transcripts")

# 3. Add useful analytic rows to generate structured comparisons
ta.add_custom_row(
    "Biggest Barrier to Profitability",
    "Summarize in one sentence the biggest barrier to profitability described in this interview (short answer, one to two sentences).\n\n{transcript}"
)
ta.add_custom_row(
    "Most Common Investment Instrument",
    "What funding instruments does the participant prefer or condone in these markets, e.g. angel, VC, etc. and debt, equity, preferred stock, etc. (short answer, one to two sentences)?\n\n{transcript}"
)
ta.add_custom_row(
    "Key VC Opportunity",
    "According to the participant, what is the most promising sector or opportunity for venture capital (short answer, one to two sentences)?\n\n{transcript}"
)

ta.add_custom_row(
    "Outlook for the Future",
    "According to the participant, is the outlook for the future positive or negative and why (short answer, one to two sentences)?\n\n{transcript}"
)
ta.add_custom_row(
    "Attitude Toward Government",
    "Summarize the participant's attitude toward the government in one sentence (short answer, one to two sentences).\n\n{transcript}"
)

# 4. Set research questions 
#ta.set_research_questions([
#    "RQ1: How do experienced local fund managers and ecosystem actors in Zambia and Kenya describe the current structure, ecosystem and dynamics of their respective investment landscapes? (Is it attractive, what are the systemic issues, and what is the outlook?)",
#    "RQ2: What operational strategies have proven effective for deploying risk capital in Zambia, and how do these compare to practices observed in Kenya’s more mature VC environment?",
#    "RQ3: Which governmental reforms are perceived by local stakeholders as the most critical to fostering a more conducive investment climate in Zambia?"
#])

# 5. Run the full 6-phase thematic analysis
ta.run()

# 6. Optionally refine codes to reduce redundancy
ta.refine_codes(method="cosine")  # 'gpt' for gpt-based merging of codes, or 'cosine' for deterministic merging

# 7. Export everything to Excel — resume if file exists
ta.to_excel("analysis.xlsx",subcolumns=["codes", "supporting_quotes"], max_quotes_per_cell=3, resume=True)
