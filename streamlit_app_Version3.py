# =========================
# 3) UPDATE build_source_location_annotation SIGNATURE + CALL
# Replace existing function definition with this
# =========================

def build_source_location_annotation(
    reference_name="",
    claim_text="",
    page_number=None,
    paragraph_number=None,
    line_range="",
    supporting_text="",
    column_number=None,
    sentence_number=None,
    column_paragraph_number=None,
):
    source_location = build_source_location_label(
        page_number=page_number,
        paragraph_number=paragraph_number,
        line_range=line_range,
        column_number=column_number,
        sentence_number=sentence_number,
        column_paragraph_number=column_paragraph_number,
    )
    claim_text = clean_text(claim_text)
    reference_name = clean_text(reference_name) or "Reference"
    supporting_text = clean_text(supporting_text)

    suggested_annotation = ""
    if source_location and claim_text:
        reference_text = f"supported by {reference_name}, {source_location.lower()}"
        suggested_annotation = f"{claim_text} — {reference_text}."

    return {
        "source_location": source_location,
        "matched_supporting_text": f'"{supporting_text}"' if supporting_text else "",
        "suggested_annotation": suggested_annotation,
    }