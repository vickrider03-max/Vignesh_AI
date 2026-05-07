import re


BROAD_INTENTS = {
    "analysis_request",
    "summarization_request",
    "technical_overview",
    "themes_request",
}


def normalize_query(text):
    return re.sub(r"\s+", " ", str(text or "").strip())


def classify_query_intent(query, previous_messages=None):
    """Classify user intent for document-intelligence routing."""
    del previous_messages
    q = normalize_query(query).lower()
    compact = re.sub(r"[^a-z0-9]+", " ", q).strip()

    if compact in {"analyze", "analyse", "analysis"} or any(
        term in q for term in ["deep analysis", "full analysis", "analyze document", "analyse document", "key insights"]
    ):
        return "analysis_request"
    if compact in {"summary", "summarize", "summarise"} or any(
        term in q for term in ["summarize", "summarise", "short summary", "main points", "recap"]
    ):
        return "summarization_request"
    if any(term in q for term in ["technical overview", "explain the architecture", "architecture", "system design"]):
        return "technical_overview"
    if any(term in q for term in ["main themes", "themes", "topics", "key topics"]):
        return "themes_request"
    if any(term in q for term in ["compare", "difference", "differences", " versus ", " vs "]):
        return "comparison_request"
    if any(term in q for term in ["metadata", "author", "created", "file type", "document type"]):
        return "metadata_request"
    if any(term in q for term in ["previous", "that", "it", "those", "follow up", "again", "same"]):
        return "followup_question"
    return "factual_query"


def to_technical_intent(document_intent):
    """Map compact routing intents onto the app's existing response intent names."""
    return {
        "analysis_request": "FULL_DOCUMENT_ANALYSIS",
        "summarization_request": "SHORT_SUMMARY",
        "technical_overview": "OVERVIEW",
        "themes_request": "OVERVIEW",
        "comparison_request": "COMPARISON",
        "metadata_request": "OVERVIEW",
        "followup_question": "QUESTION_ANSWERING",
        "factual_query": "QUESTION_ANSWERING",
    }.get(document_intent, "QUESTION_ANSWERING")


def requires_document_scope(document_intent):
    return document_intent in BROAD_INTENTS or document_intent == "metadata_request"

