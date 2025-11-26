RAG (Retrieval-Augmented Generation) workspace

Structure:
- docs/         : source documents (pdf/docx/txt). Keep originals here.
- embeddings/   : raw embedding vectors (binary files). Not tracked in git.
- index/        : vector index files (FAISS/Milvus/pgvector exports).
- meta/         : metadata JSON mapping doc IDs to sources.
- scripts/      : ingestion/ETL scripts to build embeddings and indexes.
- cache/        : temporary/cache files.

Notes:
- Large files in `embeddings/` and `index/` are ignored via `.gitignore`.
- Use `scripts/` to add ingestion steps (extract text, compute embeddings, build index).
- Optionally store documents in the DB (table `rag_documents`) instead of filesystem.

Commands:
PowerShell (create directories):
```powershell
New-Item -ItemType Directory -Path .\rag, .\rag\docs, .\rag\embeddings, .\rag\index, .\rag\meta, .\rag\scripts, .\rag\cache
```

If you want, I can:
- generate starter ingestion scripts in `scripts/` (Python), or
- create an Oracle table scaffold to store docs & embeddings.
