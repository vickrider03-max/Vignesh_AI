PROMPT_BY_INTENT = {
    "factual_query": "Answer precisely from the cited context. Prefer direct facts and short explanations.",
    "analysis_request": "Synthesize across sections. Explain purpose, architecture, themes, important details, risks, and takeaways.",
    "summarization_request": "Create a concise document-level summary using section summaries and representative evidence.",
    "comparison_request": "Compare only supported items or files. Use a criteria table, then summarize differences.",
    "followup_question": "Resolve references using conversation memory first, then answer from document context.",
    "metadata_request": "Extract file metadata, document type, structure, topics, entities, and available assets.",
    "technical_overview": "Explain the technical structure, components, interfaces, workflow, and constraints.",
    "themes_request": "Identify recurring themes/topics and explain how they appear across sections.",
}


def build_prompt(query, document_intent, context, memory="", cross_file_hints=""):
    intent_instruction = PROMPT_BY_INTENT.get(document_intent, PROMPT_BY_INTENT["factual_query"])
    return f"""You are an advanced document intelligence assistant.

Intent: {document_intent}
Task: {intent_instruction}

Rules:
- Use only the supplied context, semantic metadata, and conversation memory.
- Broad requests require synthesis across section summaries and representative pages.
- Avoid generic rejection when partial evidence exists; provide best-effort, uncertainty-aware analysis.
- Clearly say "Not specified in the provided context" only for missing details.
- Cite relevant pages, sections, sheets, tables, or diagram references.
- End with a confidence label.

Conversation memory:
{memory or "No previous conversation."}

Cross-file hints:
{cross_file_hints or "No cross-file hints."}

Context:
{context}

User question:
{query}
"""

