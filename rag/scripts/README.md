Ingest script for RAG

Files:
- `ingest.py` : main ingestion script. It:
  - reads files from `../docs`
  - extracts text (txt/docx/pdf)
  - computes embeddings (sentence-transformers by default, or OpenAI if configured)
  - saves per-document embeddings in `../embeddings/<docid>.npy`
  - builds a FAISS index in `../index/index.faiss` and mapping `mapping.json`

Requirements:
- Install the Python packages in `requirements.txt` (preferably in a venv):
  pip install -r requirements.txt

Quick start:
- Put your documents into `rag/docs`.
- Run full ingestion + index build:
  python rag/scripts/ingest.py --rebuild-index

Options:
- `--embed-only` : only compute embeddings
- `--index-only` : only build the FAISS index from existing embeddings

Notes:
- The script picks `sentence-transformers` if available; otherwise you may set `OPENAI_API_KEY` and have `openai` installed.
- FAISS installation on Windows may require special wheels; if you cannot install `faiss-cpu`, you can skip index creation and use another indexer.
- The script uses simple paragraph splitting. For production, implement smarter chunking, deduplication and metadata handling.
