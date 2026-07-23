# =========================
# 6) IN search_uploaded_article_library, include new location fields in scored rows
# Add these keys where scored/fallback candidates are built:
# =========================

# In scored.append({...}) and fallback candidate dicts, ensure:
"column_number": item.get("column_number"),
"sentence_number": item.get("sentence_number"),
"column_paragraph_number": item.get("column_paragraph_number"),