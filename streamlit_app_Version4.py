# =========================
# 4) UPDATE make_attribution_row SIGNATURE + SOURCE LOCATION CALL
# In function signature add:
#    column_number=None,
#    sentence_number=None,
#    column_paragraph_number=None,
# =========================

def make_attribution_row(
    workflow,
    claim_number,
    claim,
    source_status,
    article_title,
    source_database,
    retrieval_type,
    score,
    supporting_passage="",
    citation="",
    doi="",
    pmid="",
    url="",
    client_source="",
    recommendation="",
    page_number=None,
    paragraph_number=None,
    line_range="",
    rank=None,
    support_focus="",
    reviewer_note="",
    section_heading="",
    confidence_level="",
    source_publication_year="",
    abstract_url="",
    column_number=None,
    sentence_number=None,
    column_paragraph_number=None,
):
    source_location = build_source_location_annotation(
        reference_name=citation or article_title,
        claim_text=claim,
        page_number=page_number,
        paragraph_number=paragraph_number,
        line_range=line_range,
        supporting_text=supporting_passage,
        column_number=column_number,
        sentence_number=sentence_number,
        column_paragraph_number=column_paragraph_number,
    )

    # ... keep existing code unchanged ...

    return {
        # ... keep existing fields ...
        "column_number": column_number,
        "sentence_number": sentence_number,
        "column_paragraph_number": column_paragraph_number,
        # ... existing fields remain ...
    }