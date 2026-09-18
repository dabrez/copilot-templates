# Slide Templater

Turn a finished PowerPoint deck into a reusable template. Point it at a
`.pptx`, and it finds the things that look like they change between versions
— client names, dates, dollar figures — and swaps them for placeholders.

```
Acme Corp   →   [ORG]
John Doe    →   [PERSON]
October 14, 2023   →   [DATE]
```

Two steps: **analyze** the deck to get a JSON config of what it found, edit
that config to taste, then **template** the deck to apply it.

## Install

Requires Python 3.12.

```bash
python3 -m venv venv
./venv/bin/pip install -e .
./venv/bin/pip install -r requirements.txt   # pulls in the spaCy model
```

This installs a `slide-templater` command into the venv.

For the exact environment this was developed against, use
`requirements.lock.txt` instead.

Note that `requirements.txt` pulls in `en_core_web_trf`, a ~440MB spaCy
transformer model. It is a real dependency of the default analyze path, not
an optional extra.

## Usage

### Analyze

```bash
slide-templater analyze --input deck.pptx --output config.json
```

Writes a JSON file mapping found text to a placeholder:

```json
{
    "Acme Corp": "[ORG]",
    "October 14, 2023": "[DATE]",
    "John Doe": "[PERSON]"
}
```

This is a starting point, not an answer — read it before using it. Named
entity recognition guesses, and it will both miss things and flag things you
did not mean. Delete the lines you do not want and fix the placeholder names
you do.

`--method` picks the backend:

- `nlp` (default) — spaCy NER. Fast, deterministic, runs offline once the
  model is installed. Recognizes `DATE`, `PERSON`, `ORG`, `MONEY`, and `GPE`.
- `llm` — a local Gemma 2 2B. Downloads ~1.7GB from Hugging Face on first
  run, then caches it. Slower, but not restricted to those five categories.

### Template

```bash
slide-templater template --input deck.pptx --config config.json --output out.pptx
```

Applies the config and writes a new file. The input deck is never modified.

The config is a plain find-and-replace map, so it is useful on its own —
write one by hand to fill a template back in with real values, which is just
the analyze step run backwards:

```json
{
    "[ORG]": "Globex Inc",
    "[DATE]": "March 3, 2026"
}
```

## What it covers

Text in shapes and inside table cells, across all slides. Replacements that
span multiple formatting runs are handled — if PowerPoint internally split
`Acme Corp` into `Acme` + ` Corp` because of a stray italic, it is still
found. That case collapses the text into the first run's formatting.

Not covered: speaker notes, charts, SmartArt, images, and headers/footers.

## Notes

Replacements are applied in a single pass, so a replacement's output is never
rescanned by a later key. A config like
`{"Acme Corp": "Globex", "Globex": "WRONG"}` yields `Globex`, not `WRONG`.
Where two keys overlap, the longer one wins.

## Tests

```bash
./venv/bin/pip install -e '.[dev]'
./venv/bin/pytest
```

## Layout

```
src/slide_templater/analyzer.py       spaCy NER backend
src/slide_templater/llm_analyzer.py   Gemma backend
src/slide_templater/templater.py      find-and-replace over the .pptx
src/slide_templater/cli.py            argument parsing
tests/                                pytest suite
examples/                             sample.pptx and example configs
```

## License

Apache 2.0 — see [LICENSE](LICENSE).
