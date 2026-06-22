# Scientific Evidence Verification Framework Roadmap

## Objective
Evolve the workbench from literature discovery into a scientific evidence verification and citation validation platform with traceable evidence, rights-aware retrieval, and mandatory human oversight.

## Guiding Principles
- Evidence first: rank and report direct supporting passages before metadata-only records.
- Reproducibility: store retrieval and scoring provenance for every finding.
- Transparency: expose source type, section location, support status, and confidence.
- Compliance by design: enforce rights checks before full-text retrieval and analysis.
- Human-in-the-loop: all outputs remain draft review artifacts, never final conclusions.

## Target Architecture
- Retrieval layer: PubMed, Europe PMC, Semantic Scholar, OpenAlex, CORE, local uploads.
- Enrichment layer: EFetch abstract retrieval, section segmentation, citation context extraction.
- Verification layer: NLI classification (support/contradict/insufficient), confidence scoring.
- Semantic layer: scientific embeddings plus vector index for article/section/paragraph/sentence/citation context.
- Governance layer: rights verification, TDM restrictions, audit log, reviewer disclaimers.
- Reporting layer: reviewer-ready evidence report with full provenance.

## Priority Search Order
1. Uploaded Full Text
2. CORE Full Text
3. Europe PMC Full Text
4. PubMed Abstract
5. Semantic Scholar Metadata
6. OpenAlex Metadata

## Canonical Data Model (minimum)
Implement these entities as dataclasses or normalized dictionaries to keep ranking and reporting consistent.

### SourceRecord
- source_id
- source_type: uploaded_fulltext | core_fulltext | europepmc_fulltext | pubmed_abstract | semantic_metadata | openalex_metadata
- title, authors, journal, doi, pmid, year, url
- retrieval_status: success | partial | failed
- retrieval_reason

### EvidenceUnit
- evidence_id
- source_id
- granularity: section | paragraph | sentence | citation_context
- text
- section_name
- page_number
- paragraph_number
- citation_number

### VerificationResult
- claim_id
- evidence_id
- nli_label: SUPPORT | CONTRADICT | NOT_ENOUGH_EVIDENCE
- semantic_similarity
- evidence_quality_score
- source_authority_score
- citation_consistency_score
- support_publication_count
- confidence_level: High | Moderate | Low

### RightsAssessment
- source_id
- access_class: open_access | user_uploaded | restricted
- full_text_allowed: true | false
- tdm_restriction_detected: true | false
- ai_restriction_detected: true | false
- restriction_summary

## Feature Implementation Plan

### Phase 1: Retrieval and Evidence Grounding
Scope:
- Integrate CORE API as supplementary full-text source.
- Add PubMed EFetch for abstract text retrieval.
- Normalize retrieval provenance and status tracking.

Implementation tasks:
- Add `search_core(query, rows)` with retry/backoff and timeout.
- Add `fetch_pubmed_abstracts(pmids)` using EFetch XML/JSON parsing.
- Introduce `source_type` and `retrieval_status` fields across all search outputs.
- Refactor ranking to prioritize full-text over metadata when score ties occur.

Acceptance criteria:
- CORE queried only when higher-priority sources do not return sufficient full text.
- PubMed results include abstract text when available.
- Evidence ranking pipeline uses abstract text in scoring.
- Evidence report shows retrieval source and retrieval status for each result.

### Phase 2: Scientific PDF Parsing and Section Search
Scope:
- Integrate GROBID for structural PDF parsing.
- Index and search section-level content.

Implementation tasks:
- Add GROBID service adapter (`processFulltextDocument`).
- Map TEI output into canonical sections:
  - Abstract
  - Introduction
  - Methods
  - Results
  - Discussion
  - Conclusion
  - References
- Add search filter UI options:
  - Results only
  - Discussion only
  - Abstract only
  - Entire article
- Update ranking multipliers:
  - Results > Conclusion > Discussion > Abstract > Introduction

Acceptance criteria:
- Uploaded scientific PDFs return section-segmented text when parse succeeds.
- Evidence row includes section label and location.
- Reviewer can choose section-restricted search from UI.

### Phase 3: Citation Context Verification
Scope:
- Validate whether cited references support local claim language.

Implementation tasks:
- Use PyMuPDF to detect citation markers and extract context windows:
  - previous sentence
  - citation sentence
  - next sentence
- Resolve citation marker to reference list entry.
- Link citation entry to retrievable source metadata/full text.
- Add citation assessment labels:
  - Accurately Represented
  - Potentially Misrepresented
  - Insufficient Evidence

Acceptance criteria:
- For citation markers discovered in uploaded documents, report includes context window and resolved reference.
- Citation assessment appears in reviewer output.

### Phase 4: Claim Verification Engine (NLI)
Scope:
- Replace keyword-only support labels with evidence-based NLI classification.

Implementation tasks:
- Add pluggable NLI interface:
  - `predict_claim_evidence(claim, evidence_text) -> label, score`
- Initial model options:
  - SciFact model family
  - FEVER-adapted baseline
- Aggregate multiple evidence units per claim using max-support and contradiction-aware rules.

Acceptance criteria:
- Every claim-evidence pair has one of: SUPPORT, CONTRADICT, NOT_ENOUGH_EVIDENCE.
- Contradictory evidence is surfaced in report, not hidden by top positive score.

### Phase 5: Semantic Search Infrastructure
Scope:
- Add embedding-based retrieval for conceptual matching beyond exact terms.

Implementation tasks:
- Add scientific embedding provider abstraction.
- Start with SPECTER-compatible embeddings.
- Add Qdrant integration and collections for:
  - articles
  - sections
  - paragraphs
  - sentences
  - citation_context
- Hybrid retrieval strategy:
  - lexical score + vector score + section prior + source authority

Acceptance criteria:
- Claims retrieve semantically related evidence even without exact phrase overlap.
- Vector retrieval can be toggled and audited in output metadata.

### Phase 6: Rights and Compliance Hardening
Scope:
- Enforce retrieval boundaries based on rights signals before full-text handling.

Implementation tasks:
- Add rights decision function:
  - Open Access -> retrieve full text
  - User Uploaded -> analyze full text
  - Publisher Restricted -> metadata/abstract only
- Add mandatory warning string when full text blocked:
  - Full text was not retrieved due to copyright, licensing, or publisher restrictions.
- Add AI/TDM restriction detectors in metadata and page text checks.

Acceptance criteria:
- Full-text retrieval attempts are blocked when rights policy disallows it.
- Report explicitly logs reason when full text was not retrieved.

### Phase 7: Reviewer Transparency and Human Review Controls
Scope:
- Standardize final report for medical/regulatory traceability.

Implementation tasks:
- Add report sections:
  - Source Information
  - Evidence Information
  - Verification Information
- Include mandatory disclaimer on every export:
  - AI-assisted evidence review. Findings should be independently reviewed and approved by qualified personnel prior to use in scientific, medical, regulatory, promotional, or publication materials.

Acceptance criteria:
- Exports cannot be generated without disclaimer block.
- Each claim row contains source provenance, evidence passage, support status, confidence, and citation assessment.

## Confidence Scoring Formula (initial)
Suggested weighted score:

`confidence_raw = 0.30 * semantic_similarity + 0.20 * evidence_quality + 0.20 * source_authority + 0.15 * citation_consistency + 0.15 * support_count_normalized`

Bands:
- High: >= 0.75
- Moderate: >= 0.45 and < 0.75
- Low: < 0.45

## Reporting Contract (minimum fields)
Each evidence row in final output should include:
- Title
- Authors
- Journal
- DOI
- PMID
- Publication Year
- Source Link
- Supporting Passage
- Evidence Section
- Source Type
- Support Status
- Confidence Level
- Citation Assessment
- Retrieval Status
- Rights Decision

## Testing Strategy
- Unit tests:
  - source adapters
  - EFetch parsing
  - section mapping
  - rights decision logic
  - confidence scoring thresholds
- Integration tests:
  - claim -> retrieval -> verification -> report flow
  - fallback order validation for source priority
  - restricted-rights behavior
- Regression test set:
  - known supported claims
  - known contradicted claims
  - known misrepresented citation examples

## Operational Considerations
- External dependency resiliency: retries, timeout caps, and degraded-mode messaging.
- Cache strategy: source- and query-keyed cache with TTL and rights-aware invalidation.
- Observability: structured logs for retrieval source, model decisions, and blocked rights actions.
- Security: redact sensitive input from logs and persist only required provenance metadata.

## Delivery Milestones
- M1: CORE + EFetch + provenance tracking
- M2: GROBID section parsing + section-level search
- M3: Citation context verification and assessment labels
- M4: NLI-based claim verification and confidence scoring
- M5: Embeddings + Qdrant hybrid retrieval
- M6: Rights engine hardening + final transparency report contract

## Definition of Done for Platform Shift
The platform is considered upgraded from search to verification when all conditions are true:
- Claims are evaluated against explicit evidence units, not metadata only.
- Outputs include SUPPORT/CONTRADICT/NOT_ENOUGH_EVIDENCE labels.
- Reports include section-level evidence provenance and citation assessment.
- Rights policy gates are enforced with explicit blocked-retrieval messaging.
- Human-review disclaimer is present in all outputs.
