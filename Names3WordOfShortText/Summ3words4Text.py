#!/usr/bin/env python3
"""
Local LLM (Ollama) XML processor:
- Reads XML containing <Row> entries
- Summarizes <Note> into exactly 3 German words
- Writes summary into <NewName>

Usage:
  python Summ3words4Text.py OperationalPerformer.xml output.xml --model llama3.1:8b
"""

from __future__ import annotations

import argparse
import re
import sys
from typing import Optional
from xml.etree import ElementTree as ET

import requests


OLLAMA_URL_DEFAULT = "http://localhost:11434"
MODEL_DEFAULT = "llama3.1:8b"


def _normalize_llm_output(text: str) -> str:
    """Remove punctuation, normalize whitespace."""
    text = (text or "").strip()
    text = re.sub(r"[^\wäöüÄÖÜß\s-]", " ", text, flags=re.UNICODE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def ollama_chat(prompt: str, model: str, ollama_url: str, timeout: int = 120) -> str:
    """
    Calls Ollama /api/chat (non-streaming) and returns the assistant content.
    """
    payload = {
        "model": model,
        "stream": False,
        "messages": [
            {"role": "system", "content": "Du bist ein hilfreicher Assistent. Antworte exakt nach den Regeln."},
            {"role": "user", "content": prompt},
        ],
    }
    r = requests.post(f"{ollama_url.rstrip('/')}/api/chat", json=payload, timeout=timeout)
    r.raise_for_status()
    data = r.json()
    return (data.get("message", {}) or {}).get("content", "") or ""


def summarize_three_words_local(note_text: str, model: str, ollama_url: str) -> str:
    """
    Produce exactly 3 Englisc words (no punctuation).
    Includes a correction pass if output isn't exactly 3 words.
    """
    if not (note_text or "").strip():
        return "Kein Inhalt vorhanden"

    prompt = (
        "Fasse den folgenden Text auf Englisch in GENAU DREI WÖRTERN zusammen.\n"
        "Regeln: genau 3 Wörter; keine Satzzeichen; keine Anführungszeichen; keine Emojis; "
        "keine Zeilenumbrüche; gib NUR die drei Wörter aus.\n\n"
        f"TEXT:\n{note_text.strip()}\n"
    )
    out = _normalize_llm_output(ollama_chat(prompt, model=model, ollama_url=ollama_url))
    words = out.split()
    if len(words) == 3:
        return " ".join(words)

    fix_prompt = (
        "Korrigiere auf GENAU 3 Wörter (Englisch). Keine Satzzeichen. "
        "Gib NUR die drei Wörter aus.\n"
        f"Aktuelle Ausgabe: {out}"
    )
    out2 = _normalize_llm_output(ollama_chat(fix_prompt, model=model, ollama_url=ollama_url))
    words2 = out2.split()
    if len(words2) == 3:
        return " ".join(words2)

    # Last resort: truncate/pad
    final_words = (words2[:3] + ["…", "…", "…"])[:3]
    return " ".join(final_words)


def process_xml(input_path: str, output_path: str, model: str, ollama_url: str) -> None:
    tree = ET.parse(input_path)
    root = tree.getroot()

    rows = root.findall(".//Row")
    for row in rows:
        note_el = row.find("Note")
        note_text = note_el.text if (note_el is not None and note_el.text) else ""

        summary = summarize_three_words_local(note_text, model=model, ollama_url=ollama_url)

        newname_el = row.find("NewName")
        if newname_el is None:
            newname_el = ET.SubElement(row, "NewName")
        newname_el.text = summary

    # Pretty print (Python 3.9+)
    try:
        ET.indent(tree, space="  ")
    except Exception:
        pass

    tree.write(output_path, encoding="utf-8", xml_declaration=True)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("input_xml")
    ap.add_argument("output_xml")
    ap.add_argument("--model", default=MODEL_DEFAULT, help="Ollama model name, e.g. llama3.1:8b")
    ap.add_argument("--ollama-url", default=OLLAMA_URL_DEFAULT, help="Ollama base URL (default: http://localhost:11434)")
    args = ap.parse_args()

    try:
        process_xml(args.input_xml, args.output_xml, model=args.model, ollama_url=args.ollama_url)
        print(f"Done. Wrote updated XML to: {args.output_xml}")
        return 0
    except requests.RequestException as e:
        print(f"HTTP error talking to Ollama: {e}", file=sys.stderr)
        return 2
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
