import unittest

from reference_source_claims import select_reference_source_claims

class ReferenceSourceClaimSplittingTests(unittest.TestCase):
    def test_multi_sentence_claim_box_input_remains_single_claim(self):
        split_calls = []

        def split_into_claims_impl(content, max_claims=10):
            split_calls.append((content, max_claims))
            return ["unexpected_claim_1", "unexpected_claim_2"]

        pasted_text = "Sentence one. Sentence two with more context. Sentence three."

        claims = select_reference_source_claims(
            content=pasted_text,
            max_claims=5,
            split_claims=False,
            split_into_claims=split_into_claims_impl,
        )

        self.assertEqual(split_calls, [])
        self.assertEqual(claims, [pasted_text])

    def test_multiline_claim_box_input_remains_single_claim(self):
        split_calls = []

        def split_into_claims_impl(content, max_claims=10):
            split_calls.append((content, max_claims))
            return ["unexpected_claim"]

        pasted_text = "Line one\nLine two\nLine three"

        claims = select_reference_source_claims(
            content=pasted_text,
            max_claims=5,
            split_claims=False,
            split_into_claims=split_into_claims_impl,
        )

        self.assertEqual(split_calls, [])
        self.assertEqual(claims, [pasted_text])

    def test_non_claim_box_flow_keeps_prior_split_behavior(self):
        split_calls = []

        def split_into_claims_impl(content, max_claims=10):
            split_calls.append((content, max_claims))
            return ["First split claim", "Second split claim"]

        claims = select_reference_source_claims(
            content="First sentence. Second sentence.",
            max_claims=5,
            split_claims=True,
            split_into_claims=split_into_claims_impl,
        )

        self.assertEqual(split_calls, [("First sentence. Second sentence.", 5)])
        self.assertEqual(claims, ["First split claim", "Second split claim"])


if __name__ == "__main__":
    unittest.main()
