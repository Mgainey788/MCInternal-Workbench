# =========================
# 1) ADD THIS HELPER BLOCK
# Place near other text helpers (e.g., after split_sentences / clean helpers)
# =========================

def split_text_into_paragraphs(text):
    """
    Robust paragraph splitter for extracted PDF text.
    """
    raw = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    parts = re.split(r"\n{2,}", raw)
    out = []
    for p in parts:
        p = clean_text(p)
        if len(p) >= 20:
            out.append(p)
    if not out:
        # fallback single block
        one = clean_text(raw)
        if one:
            out = [one]
    return out


def split_paragraph_into_sentences(paragraph):
    """
    Sentence splitter used for sentence-level location annotations.
    """
    para = clean_text(paragraph)
    if not para:
        return []
    sents = [clean_text(s) for s in re.split(r"(?<=[.!?])\s+", para) if clean_text(s)]
    return sents if sents else [para]


def infer_columns_from_page_text(page_text):
    """
    Heuristic column detection from extracted page text.
    NOTE: With pypdf plain text extraction, this is best-effort only.
    Priority:
      1) explicit column markers if present
      2) fallback split by long line breaks / paragraph mass
    Returns list of column text blocks [col1_text, col2_text?]
    """
    txt = page_text or ""
    txt_clean = clean_text(txt)
    if not txt_clean:
        return []

    # Explicit marker support (if upstream ever provides markers)
    # e.g. [[COLUMN:1]] ... [[COLUMN:2]] ...
    if "[[COLUMN:" in txt:
        chunks = re.split(r"(?=\[\[COLUMN:\d+\]\])", txt)
        cols = []
        for chunk in chunks:
            chunk = re.sub(r"\[\[COLUMN:\d+\]\]", " ", chunk)
            chunk = clean_text(chunk)
            if chunk:
                cols.append(chunk)
        if cols:
            return cols

    # Heuristic split:
    # If many paragraphs and long body, split into 2 columns by paragraph midpoint.
    paras = split_text_into_paragraphs(txt)
    if len(paras) >= 8:
        mid = len(paras) // 2
        col1 = clean_text(" ".join(paras[:mid]))
        col2 = clean_text(" ".join(paras[mid:]))
        if col1 and col2:
            return [col1, col2]

    # fallback single column
    return [txt_clean]