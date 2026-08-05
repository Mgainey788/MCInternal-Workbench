import builtins
import importlib
import sys


def test_optional_sentence_transformers_import_failure_is_handled(monkeypatch):
    real_import = builtins.__import__

    def fake_import(name, globals=None, locals=None, fromlist=(), level=0):
        if name == "sentence_transformers":
            raise RuntimeError("simulated import failure")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    sys.modules.pop("streamlit_app", None)

    module = importlib.import_module("streamlit_app")

    assert module.SentenceTransformer is None
    assert module._SENTENCE_TRANSFORMERS_AVAILABLE is False
