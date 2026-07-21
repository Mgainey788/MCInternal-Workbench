def select_reference_source_claims(content, max_claims, split_claims, split_into_claims):
    """Return claims for reference-source search, optionally splitting or preserving full input."""
    normalized_content = (content or "").strip()

    if split_claims:
        claims = split_into_claims(normalized_content, max_claims=max_claims)
        if not claims and normalized_content:
            claims = [normalized_content]
        return claims

    return [normalized_content] if normalized_content else []
