# =========================
# 8) UPDATE render_professional_rows where Source Location is built
# Replace existing lambda call in rank_table["Source Location"] with:
# =========================

rank_table["Source Location"] = rank_table.apply(
    lambda row: build_source_location_label(
        row.get("page_number"),
        row.get("paragraph_number"),
        row.get("line_range", ""),
        column_number=row.get("column_number"),
        sentence_number=row.get("sentence_number"),
        column_paragraph_number=row.get("column_paragraph_number"),
    ) or "-",
    axis=1,
)