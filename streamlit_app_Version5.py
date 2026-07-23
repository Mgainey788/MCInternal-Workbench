# =========================
# 5) REPLACE split_article_into_passages with column/sentence aware version
# =========================

def split_article_into_passages(article_text, source_name=""):
    raw_text = article_text or ""
    if not clean_text(raw_text):
        return []

    passages = []
    running_passage_number = 0
    source_publication_year = infer_publication_year_from_document_text(raw_text, source_name=source_name)

    # Keep page labels from PDF extraction
    page_chunks = re.split(r"(?=\[\[PDF_PAGE_LABEL:[^\]]+\]\])|(?=\bPage\s+\d+\s*:)", raw_text)
    labeled_chunks = []

    for chunk in page_chunks:
        page_label_match = re.match(r"\s*\[\[PDF_PAGE_LABEL:([^\]]+)\]\]\s*(.*)", chunk, flags=re.IGNORECASE | re.DOTALL)
        if page_label_match:
            labeled_chunks.append((clean_text(page_label_match.group(1)), page_label_match.group(2)))
            continue

        match = re.match(r"\s*Page\s+(\d+)\s*:\s*(.*)", chunk, flags=re.IGNORECASE | re.DOTALL)
        if match:
            labeled_chunks.append((match.group(1), match.group(2)))

    if not labeled_chunks:
        labeled_chunks = [(None, raw_text)]

    for page_number, chunk_text in labeled_chunks:
        cleaned_chunk = clean_text(chunk_text)
        if not cleaned_chunk:
            continue

        cleaned_chunk = remove_abstract_language(cleaned_chunk)
        if not cleaned_chunk:
            continue

        columns = infer_columns_from_page_text(cleaned_chunk)
        if not columns:
            columns = [cleaned_chunk]

        page_sentence_counter = 0  # sentence index within page

        for col_idx, col_text in enumerate(columns, start=1):
            col_section_label = classify_chunk_section(col_text)
            col_paragraph_counter = 0

            paragraphs = split_text_into_paragraphs(col_text)
            for paragraph in paragraphs:
                col_paragraph_counter += 1
                paragraph = clean_text(paragraph)
                if len(paragraph) < 30:
                    continue

                sentences = split_paragraph_into_sentences(paragraph)
                for sentence in sentences:
                    sentence = clean_text(sentence)
                    if len(sentence) < 40:
                        continue

                    running_passage_number += 1
                    page_sentence_counter += 1

                    section_label = col_section_label if col_section_label != "body" else classify_passage_section(sentence)

                    page_i = _annotation_int(page_number)
                    if page_i == 1 and col_paragraph_counter <= 8 and section_label == "body":
                        section_label = "introduction"

                    passages.append({
                        "source_name": source_name,
                        "source_publication_year": source_publication_year,
                        "passage_number": running_passage_number,
                        "page_paragraph_number": col_paragraph_counter,  # kept for backward compatibility
                        "column_paragraph_number": col_paragraph_counter,
                        "sentence_number": page_sentence_counter,
                        "column_number": col_idx if len(columns) > 1 else 1,
                        "page_number": page_number,
                        "section_label": section_label,
                        "passage": sentence,
                    })

    return passages