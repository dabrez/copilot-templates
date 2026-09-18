"""Tests for the LLM backend's output handling.

The model itself is not exercised here -- it is a 1.6GB download and is
nondeterministic. What is worth pinning is the code around it: the prompt
assembly and the parsing of whatever the model returns.
"""

import json
import sys
import types

import pytest

llama_cpp = pytest.importorskip(
    "llama_cpp", reason="requires the [llm] extra"
)

from slide_templater import llm_analyzer


class FakeLlama:
    """Stands in for Llama, returning a canned completion."""

    def __init__(self, text):
        self._text = text
        self.prompt = None

    def __call__(self, prompt, **kwargs):
        self.prompt = prompt
        return {"choices": [{"text": self._text}]}


@pytest.fixture
def fake_model(monkeypatch, tmp_path):
    """Patch out the download and the model, yielding a setter for the reply."""
    holder = {}

    monkeypatch.setattr(
        llm_analyzer, "hf_hub_download", lambda **kw: str(tmp_path / "m.gguf")
    )

    def install(text):
        holder["llm"] = FakeLlama(text)
        monkeypatch.setattr(llm_analyzer, "Llama", lambda **kw: holder["llm"])
        return holder["llm"]

    return install


def test_parses_a_well_formed_reply(fake_model, tmp_path, repo_root):
    # The prompt pre-seeds an opening brace, so the model's reply omits it.
    fake_model('"Acme Corp": "[ORG]", "John Doe": "[PERSON]"}')
    out = tmp_path / "c.json"

    llm_analyzer.analyze_presentation_llm(
        repo_root / "examples" / "sample.pptx", out
    )

    assert json.loads(out.read_text()) == {
        "Acme Corp": "[ORG]",
        "John Doe": "[PERSON]",
    }


def test_strips_trailing_prose_after_the_json(fake_model, tmp_path, repo_root):
    """Models often keep talking after the closing brace."""
    fake_model('"Acme Corp": "[ORG]"}\n\nI hope this helps!')
    out = tmp_path / "c.json"

    llm_analyzer.analyze_presentation_llm(
        repo_root / "examples" / "sample.pptx", out
    )

    assert json.loads(out.read_text()) == {"Acme Corp": "[ORG]"}


def test_unparseable_reply_raises_runtime_error(fake_model, tmp_path, repo_root):
    fake_model("this is not json at all")

    with pytest.raises(RuntimeError, match="valid JSON"):
        llm_analyzer.analyze_presentation_llm(
            repo_root / "examples" / "sample.pptx", tmp_path / "c.json"
        )


def test_truncated_reply_raises_rather_than_writing_partial_output(
    fake_model, tmp_path, repo_root
):
    """A reply cut off by max_tokens must not yield a half-written config.

    The recovery step trims to the last '}', so a truncation that happens to
    leave one behind can still parse -- but it must never silently produce a
    file from a reply that does not parse at all.
    """
    fake_model('"Acme Corp": "[ORG]", "October 14, 2023": "[DA')
    out = tmp_path / "c.json"

    with pytest.raises(RuntimeError):
        llm_analyzer.analyze_presentation_llm(
            repo_root / "examples" / "sample.pptx", out
        )
    assert not out.exists()


def test_missing_input_raises_before_loading_the_model(tmp_path):
    with pytest.raises(FileNotFoundError):
        llm_analyzer.analyze_presentation_llm(
            tmp_path / "nope.pptx", tmp_path / "c.json"
        )


def test_prompt_carries_the_deck_text(fake_model, tmp_path, repo_root):
    llm = fake_model('"Acme Corp": "[ORG]"}')

    llm_analyzer.analyze_presentation_llm(
        repo_root / "examples" / "sample.pptx", tmp_path / "c.json"
    )

    assert "Acme Corp" in llm.prompt
    assert llm.prompt.rstrip().endswith("{")
