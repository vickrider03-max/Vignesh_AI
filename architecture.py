TARGET_FOLDER_STRUCTURE = """
document_intelligence/
  query_routing.py        # intent detection and broad-vs-focused routing
  document_processing.py  # semantic metadata, summaries, topics, domains
  retrieval.py            # hybrid retrieval and context expansion
  reranking.py            # lexical/cross-encoder reranking
  prompts.py              # task-specific prompt templates
  fallback.py             # best-effort uncertainty-aware fallback answers
  memory.py               # conversation and file memory contracts
  storage.py              # persistence boundaries for file brains/vector stores
  handlers.py             # request handlers per intent
  utils.py                # shared pure helpers
"""


DATA_FLOW = """
UPLOAD
  -> file type detection
  -> page/slide/sheet extraction
  -> file brain build
       -> page index
       -> section summaries
       -> document summary
       -> topics/entities/domains
       -> facts/tables/diagrams
       -> suggested questions
  -> global memory registry
  -> query intent classification
  -> hybrid retrieval
       -> document-level context for broad requests
       -> page/table/diagram context for focused requests
       -> BM25 + metadata + lazy selected-chunk embeddings
  -> reranking
  -> specialized prompt
  -> best-effort fallback if context is thin
  -> response with citations and confidence
"""

