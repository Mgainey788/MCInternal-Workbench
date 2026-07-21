import ast
import pathlib
import unittest


REPO_ROOT = pathlib.Path(__file__).resolve().parents[1]
STREAMLIT_APP_PATH = REPO_ROOT / "streamlit_app.py"


def load_function(function_name, globals_overrides):
    source = STREAMLIT_APP_PATH.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(STREAMLIT_APP_PATH))

    function_node = next(
        node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == function_name
    )
    module = ast.Module(body=[function_node], type_ignores=[])
    namespace = dict(globals_overrides)
    exec(compile(module, str(STREAMLIT_APP_PATH), "exec"), namespace)
    return namespace[function_name]


class ReferenceSourceClaimSplittingTests(unittest.TestCase):
    def _build_workflow(self, split_into_claims_impl):
        search_calls = []

        def search_for_true_source(*, claim, **kwargs):
            search_calls.append(claim)
            return []

        workflow = load_function(
            "run_reference_source_workflow",
            {
                "run_citation_qa_resolver": lambda **kwargs: {"matched": False},
                "citation_qa_to_log_row": lambda *args, **kwargs: {},
                "split_into_claims": split_into_claims_impl,
                "classify_statement_type": lambda claim: "General claim",
                "resolve_exact_source_quote": lambda *args, **kwargs: {"matched": False},
                "make_attribution_row": lambda **kwargs: kwargs,
                "search_for_true_source": search_for_true_source,
            },
        )
        return workflow, search_calls

    def test_multi_sentence_claim_box_input_remains_single_claim(self):
        split_calls = []

        def split_into_claims_impl(content, max_claims=10):
            split_calls.append((content, max_claims))
            return ["SHOULD NOT BE USED", "SHOULD NOT BE USED EITHER"]

        workflow, search_calls = self._build_workflow(split_into_claims_impl)
        pasted_text = "Sentence one. Sentence two with more context. Sentence three."

        workflow(content=pasted_text, max_claims=5, split_claims=False)

        self.assertEqual(split_calls, [])
        self.assertEqual(search_calls, [pasted_text])

    def test_multiline_claim_box_input_remains_single_claim(self):
        split_calls = []

        def split_into_claims_impl(content, max_claims=10):
            split_calls.append((content, max_claims))
            return ["SHOULD NOT BE USED"]

        workflow, search_calls = self._build_workflow(split_into_claims_impl)
        pasted_text = "Line one\nLine two\nLine three"

        workflow(content=pasted_text, max_claims=5, split_claims=False)

        self.assertEqual(split_calls, [])
        self.assertEqual(search_calls, [pasted_text])

    def test_non_claim_box_flow_keeps_prior_split_behavior(self):
        split_calls = []

        def split_into_claims_impl(content, max_claims=10):
            split_calls.append((content, max_claims))
            return ["First split claim", "Second split claim"]

        workflow, search_calls = self._build_workflow(split_into_claims_impl)

        workflow(content="First sentence. Second sentence.", max_claims=5)

        self.assertEqual(split_calls, [("First sentence. Second sentence.", 5)])
        self.assertEqual(search_calls, ["First split claim", "Second split claim"])


if __name__ == "__main__":
    unittest.main()
