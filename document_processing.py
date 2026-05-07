import re
from collections import Counter


STOPWORDS = {
    "the", "and", "for", "with", "that", "this", "from", "are", "was", "were", "into", "your",
    "have", "has", "had", "not", "but", "you", "all", "can", "will", "using", "used", "how",
    "what", "when", "where", "which", "while", "about", "their", "there", "page", "pages",
    "document", "content", "table", "figure", "section", "information", "data",
}

DOMAIN_TERMS = {
    "automotive": {"can", "capl", "canoe", "canalyzer", "ecu", "vehicle", "lin", "flexray"},
    "software": {"api", "code", "function", "class", "module", "database", "server", "client"},
    "networking": {"ethernet", "tcp", "ip", "network", "gateway", "protocol", "bus"},
    "electronics": {"pin", "connector", "signal", "voltage", "current", "channel", "wiring"},
    "business": {"revenue", "customer", "market", "sales", "cost", "risk", "strategy"},
    "testing": {"test", "validation", "requirement", "result", "pass", "fail", "coverage"},
}


def clean_line(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def tokenize(text):
    return [
        token
        for token in re.findall(r"[A-Za-z][A-Za-z0-9_-]{2,}", str(text or "").lower())
        if token not in STOPWORDS
    ]


def meaningful_sentences(text, limit=8):
    sentences = []
    for sentence in re.split(r"(?<=[.!?])\s+|\n+", str(text or "")):
        sentence = clean_line(sentence)
        if 45 <= len(sentence) <= 320:
            sentences.append(sentence)
        if len(sentences) >= limit:
            break
    return sentences


def summarize_page(page_record, max_sentences=3):
    text = page_record.get("text", "")
    sentences = meaningful_sentences(text, limit=max_sentences)
    if sentences:
        summary = " ".join(sentences)
    else:
        summary = clean_line(text)[:420]
    return {
        "page": page_record.get("page"),
        "section": page_record.get("section") or f"Page {page_record.get('page')}",
        "summary": summary,
        "keywords": [word for word, _ in Counter(tokenize(text)).most_common(12)],
    }


def extract_topics(pages, tables=None, limit=14):
    counter = Counter()
    for page in pages or []:
        counter.update(tokenize(page.get("text", "")))
    for table in tables or []:
        counter.update(tokenize(" ".join(str(h) for h in table.get("headers", []))))
    return [word for word, _ in counter.most_common(limit)]


def detect_technical_domains(topics, pages):
    topic_set = set(topics or [])
    combined_sample = " ".join(str(page.get("text", ""))[:1200].lower() for page in (pages or [])[:12])
    domains = []
    for domain, terms in DOMAIN_TERMS.items():
        score = len(topic_set.intersection(terms)) + sum(1 for term in terms if term in combined_sample)
        if score:
            domains.append({"domain": domain, "score": score})
    domains.sort(key=lambda item: item["score"], reverse=True)
    return domains[:5]


def build_master_summary(section_summaries, topics, tables=None, diagrams=None):
    useful = [
        item.get("summary", "")
        for item in section_summaries or []
        if item.get("summary")
    ][:6]
    if useful:
        overview = " ".join(useful)
    else:
        overview = "No meaningful document narrative was extracted."
    topic_text = ", ".join(topics[:10]) if topics else "No dominant topics detected"
    return (
        f"Document-level summary: {overview[:1800]}\n\n"
        f"Dominant topics: {topic_text}.\n"
        f"Structured tables indexed: {len(tables or [])}. Diagram references indexed: {len(diagrams or [])}."
    )


def suggested_questions(topics, domains, has_tables=False, has_diagrams=False):
    topic = topics[0] if topics else "this document"
    suggestions = [
        "Summarize the document",
        "Give key insights",
        f"Explain {topic}",
        "Provide a technical overview",
    ]
    if domains:
        suggestions.append(f"Explain the {domains[0]['domain']} context")
    if has_tables:
        suggestions.append("Extract the relevant tables")
    if has_diagrams:
        suggestions.append("Explain the diagrams or figures")
    return suggestions[:6]


def build_hierarchy(section_summaries):
    return [
        {
            "level": "section",
            "page": item.get("page"),
            "title": item.get("section"),
            "keywords": item.get("keywords", []),
        }
        for item in section_summaries or []
    ]


def build_semantic_metadata(file_name, file_type, pages, tables=None, diagrams=None, entities=None):
    """Create document-level understanding artifacts at upload/file-brain time."""
    section_summaries = [summarize_page(page) for page in pages or []]
    topics = extract_topics(pages, tables=tables)
    domains = detect_technical_domains(topics, pages)
    return {
        "metadata": {
            "file_name": file_name,
            "file_type": file_type,
            "page_or_section_count": len(pages or []),
            "table_count": len(tables or []),
            "diagram_count": len(diagrams or []),
        },
        "section_summaries": section_summaries,
        "document_summary": build_master_summary(section_summaries, topics, tables=tables, diagrams=diagrams),
        "topics": topics,
        "entities": list(entities or [])[:500],
        "technical_domains": domains,
        "suggested_questions": suggested_questions(topics, domains, bool(tables), bool(diagrams)),
        "hierarchy": build_hierarchy(section_summaries),
    }

