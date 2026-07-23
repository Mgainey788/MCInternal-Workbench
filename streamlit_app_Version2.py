# =========================
# 2) UPDATE build_source_location_label SIGNATURE + BODY
# Replace existing function:
# def build_source_location_label(page_number=None, paragraph_number=None, line_range=""):
# =========================

def build_source_location_label(
    page_number=None,
    paragraph_number=None,
    line_range="",
    column_number=None,
    sentence_number=None,
    column_paragraph_number=None,
):
    page_text = _annotation_page_value(page_number)
    para_i = _annotation_int(paragraph_number)
    col_i = _annotation_int(column_number)
    col_para_i = _annotation_int(column_paragraph_number)
    sent_i = _annotation_int(sentence_number)

    if page_text is None:
        return ""

    parts = [f"Page {page_text}"]

    if col_i is not None:
        parts.append(f"Column {col_i}")

    # Prefer column paragraph if available, else legacy paragraph_number
    if col_para_i is not None:
        parts.append(f"Paragraph {col_para_i}")
    elif para_i is not None:
        parts.append(f"Paragraph {para_i}")

    if sent_i is not None:
        parts.append(f"Sentence {sent_i}")

    label = ", ".join(parts)

    line_text = clean_text(line_range)
    if line_text:
        label += f", Line(s) {line_text}"
    return label