#!/usr/bin/env python3
"""
Simple RAG ingestion script
- Extracts text from files in ../docs
- Computes embeddings (sentence-transformers or OpenAI)
- Stores embeddings in ../embeddings (one .npy per document)
- Builds a FAISS index in ../index

Usage:
  python ingest.py [--rebuild-index] [--embed-only] [--index-only]

Make sure to run in the project root (where `rag/` is located).
"""
import os
import sys
import argparse
import json
from pathlib import Path
from hashlib import sha1

PROJECT_ROOT = Path(__file__).resolve().parents[2]
RAG_DIR = PROJECT_ROOT / 'rag'
DOCS_DIR = RAG_DIR / 'docs'
EMB_DIR = RAG_DIR / 'embeddings'
INDEX_DIR = RAG_DIR / 'index'
META_DIR = RAG_DIR / 'meta'

for d in (EMB_DIR, INDEX_DIR, META_DIR):
    d.mkdir(parents=True, exist_ok=True)

# try optional libs
try:
    from sentence_transformers import SentenceTransformer
    _HAS_ST = True
except Exception:
    _HAS_ST = False

try:
    import openai
    _HAS_OPENAI = True
except Exception:
    _HAS_OPENAI = False

try:
    import faiss
    _HAS_FAISS = True
except Exception:
    _HAS_FAISS = False

import numpy as np
from tqdm import tqdm

# text extractors
def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == '.txt':
        return path.read_text(encoding='utf-8', errors='ignore')
    if suffix == '.docx':
        try:
            from docx import Document
        except Exception:
            raise RuntimeError('python-docx not installed (pip install python-docx)')
        doc = Document(path)
        return '\n'.join(p.text for p in doc.paragraphs)
    if suffix == '.pdf':
        try:
            from pdfminer.high_level import extract_text
        except Exception:
            raise RuntimeError('pdfminer.six not installed (pip install pdfminer.six)')
        return extract_text(str(path))
    # fallback: try reading as text
    try:
        return path.read_text(encoding='utf-8', errors='ignore')
    except Exception:
        return ''

# embedding provider
class Embedder:
    def __init__(self, provider='st'):
        self.provider = provider
        if provider == 'st':
            if not _HAS_ST:
                raise RuntimeError('sentence-transformers not available')
            self.model = SentenceTransformer('all-MiniLM-L6-v2')
        elif provider == 'openai':
            if not _HAS_OPENAI:
                raise RuntimeError('openai package not available')
            if 'OPENAI_API_KEY' not in os.environ:
                raise RuntimeError('OPENAI_API_KEY not set')
            openai.api_key = os.environ['OPENAI_API_KEY']
        else:
            raise ValueError('Unknown provider')

    def embed(self, texts):
        if self.provider == 'st':
            emb = self.model.encode(texts, show_progress_bar=False)
            return np.array(emb, dtype=np.float32)
        else:
            # openai embeddings
            res = openai.Embedding.create(input=texts, model='text-embedding-3-small')
            arr = [r['embedding'] for r in res['data']]
            return np.array(arr, dtype=np.float32)

# helpers
def doc_id_for_path(path: Path) -> str:
    # stable id based on file path
    h = sha1(str(path).encode('utf-8')).hexdigest()
    return h

def ingest_all(args):
    provider = 'st' if _HAS_ST else ('openai' if _HAS_OPENAI else None)
    if provider is None and not args.index_only:
        print('No embedding provider available. Install sentence-transformers or openai and set API key.')
        sys.exit(1)
    embedder = None
    if not args.index_only:
        embedder = Embedder(provider=provider)

    meta = {}
    files = sorted([p for p in DOCS_DIR.glob('*') if p.is_file()])
    for path in tqdm(files, desc='Documents'):
        doc_id = doc_id_for_path(path)
        text = extract_text(path)
        if not text or len(text.strip())==0:
            print(f'skipping empty: {path.name}')
            continue
        # simple chunking: split into paragraphs
        paragraphs = [p.strip() for p in text.split('\n') if p.strip()]
        if not paragraphs:
            continue
        if not args.index_only:
            embeddings = embedder.embed(paragraphs)
            # save embeddings per document
            np.save(EMB_DIR / f'{doc_id}.npy', embeddings)
        meta[doc_id] = {
            'file': str(path.name),
            'chunks': len(paragraphs)
        }
    # write meta
    with open(META_DIR / 'docs.json', 'w', encoding='utf8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    print('Ingestion complete. docs:', len(meta))

    if args.rebuild_index or args.embed_only is False:
        build_index()

def build_index():
    if not _HAS_FAISS:
        print('faiss not available; install faiss-cpu to build index')
        return
    # load all embeddings and build flat index
    meta_path = META_DIR / 'docs.json'
    if not meta_path.exists():
        print('no metadata found - run ingestion first')
        return
    with open(meta_path, 'r', encoding='utf8') as f:
        meta = json.load(f)
    all_vecs = []
    mapping = []  # [(doc_id, chunk_index, filename)]
    for doc_id, info in meta.items():
        emb_file = EMB_DIR / f'{doc_id}.npy'
        if not emb_file.exists():
            print('missing embeddings for', doc_id)
            continue
        arr = np.load(emb_file)
        for i in range(arr.shape[0]):
            all_vecs.append(arr[i])
            mapping.append({'doc_id': doc_id, 'chunk': i, 'file': info['file']})
    if not all_vecs:
        print('no vectors to index')
        return
    X = np.vstack(all_vecs).astype('float32')
    d = X.shape[1]
    index = faiss.IndexFlatL2(d)
    index.add(X)
    faiss.write_index(index, str(INDEX_DIR / 'index.faiss'))
    # write mapping
    with open(INDEX_DIR / 'mapping.json', 'w', encoding='utf8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    print('Index built: vectors=', X.shape[0], 'dim=', d)

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--rebuild-index', action='store_true', help='rebuild FAISS index after ingest')
    parser.add_argument('--embed-only', action='store_true', help='only create embeddings, do not build index')
    parser.add_argument('--index-only', action='store_true', help='only build index from existing embeddings')
    args = parser.parse_args()

    if args.index_only:
        build_index()
    else:
        ingest_all(args)
