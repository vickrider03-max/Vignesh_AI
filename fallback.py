import re


def _clean(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def build_best_effort_response(query, document_intent, brains, sources_text):
    """Graceful fallback for broad questions when retrieval/LLM coverage is thin."""
    blocks = []
    for file_name, brain in (brains or {}).items():
        semantic = brain.get("semantic_metadata", {}) or {}
        summary = semantic.get("document_summary") or ""
        topics = semantic.get("topics") or brain.get("entities", [])[:10]
        sections = semantic.get("section_summaries") or []
        domains = semantic.get("technical_domains") or []

        if not summary and not sections and not topics:
            continue

        blocks.append(f"### {file_name}")
        if summary:
            blocks.append(_clean(summary)[:1200])
        if topics:
            blocks.append("**Key topics:** " + ", ".join(str(topic) for topic in topics[:10]))
        if domains:
            blocks.append("**Likely technical domains:** " + ", ".join(item.get("domain", "") for item in domains[:4]))
        if sections:
            section_lines = []
            for item in sections[:6]:
                label = item.get("section") or f"Page {item.get('page')}"
                section_lines.append(f"- {label}: {_clean(item.get('summary', ''))[:220]}")
            blocks.append("**Relevant sections:**\n" + "\n".join(section_lines))

    if not blocks:
        return (
            "Answer:\nI could not build enough document-level context from the selected files. "
            "Try re-uploading the document or ask for a specific page, section, table, or keyword.\n\n"
            f"Sources:\n{sources_text or '- No sources found'}"
        )

    caveat = (
        "This is a best-effort document-level synthesis from extracted summaries, topics, entities, "
        "tables, and diagram metadata. Details not represented in those artifacts are not assumed."
    )
    return (
        f"Answer:\n{caveat}\n\n"
        + "\n\n".join(blocks)
        + f"\n\nSources:\n{sources_text or '- File brain semantic metadata'}"
    )

