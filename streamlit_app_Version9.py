source_location = loc.get("source_location_display") or build_source_location_label(
    loc.get("page_number"),
    loc.get("paragraph_number"),
    loc.get("line_range", ""),
    column_number=loc.get("column_number"),
    sentence_number=loc.get("sentence_number"),
    column_paragraph_number=loc.get("column_paragraph_number"),
)