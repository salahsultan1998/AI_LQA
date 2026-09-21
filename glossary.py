"""
Glossary management module for fetching and matching game terms.
"""

from typing import Any
from config import Settings, get_supabase, ConfigError

# Map common Excel column names to your Supabase English language names
LANGUAGE_MAP = {
    "英文": "english",
    "英语": "english",
    "en": "english",
    "en-us": "english",
    "中文": "chinese",
    "简体中文": "chinese",
    "繁体中文": "chinese",
    "zh": "chinese",
    "日文": "japanese",
    "日语": "japanese",
    "ja": "japanese",
    "韩文": "korean",
    "韩语": "korean",
    "ko": "korean",
    "法文": "french",
    "法语": "french",
    "fr": "french",
    "德文": "german",
    "德语": "german",
    "de": "german",
    "西班牙文": "spanish",
    "西班牙语": "spanish",
    "es": "spanish",
    "俄文": "russian",
    "俄语": "russian",
    "ru": "russian",
    "阿拉伯文": "arabic",
    "阿拉伯语": "arabic",
    "ar": "arabic"
}

def fetch_glossary(
    settings: Settings, game_id: str, language: str
) -> list[dict[str, Any]]:
    """
    Load all glossary rows for a game and language from Supabase.
    """
    if not game_id or not language:
        return []

    try:
        client = get_supabase(settings)
    except ConfigError as e:
        print(f"Skipping glossary fetch: {e}")
        return []

    game = game_id.strip().lower()
    
    # 1. Clean the incoming Excel column name
    raw_lang = language.strip().lower()
    # 2. Look it up in the map. If not found, fall back to the raw string.
    lang = LANGUAGE_MAP.get(raw_lang, raw_lang)
    
    print(f"DEBUG: Mapped Excel column '{language}' to Supabase language '{lang}'")

    rows: list[dict[str, Any]] = []
    start = 0
    page_size = settings.glossary_page_size

    # Paginate past Supabase 1000-row default limit
    while True:
        end = start + page_size - 1
        try:
            response = (
                client.table("glossaries")
                .select("source_term,target_term")
                .eq("game_id", game)
                .eq("target_language", lang)
                .range(start, end)
                .execute()
            )
            page = response.data or []
            rows.extend(page)

            if len(page) < page_size:
                break
            start += page_size
        except Exception as err:
            print(f"Error querying Supabase glossaries: {err}")
            break

    # Clean and deduplicate results
    cleaned: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()

    for row in rows:
        source_term = str(row.get("source_term") or "").strip()
        target_term = str(row.get("target_term") or "").strip()

        if not source_term:
            continue

        key = (source_term, target_term)
        if key in seen:
            continue

        seen.add(key)
        cleaned.append({"source_term": source_term, "target_term": target_term})

    return cleaned


def match_terms(terms: list[dict[str, str]], text: str) -> list[dict[str, str]]:
    """
    Longest-first matching algorithm.
    Ensures longer terms take precedence so shorter sub-terms inside them are not double-matched.
    """
    if not text or not terms:
        return []

    # Sort terms by length of source term in descending order
    ranked = sorted(terms, key=lambda item: len(item["source_term"]), reverse=True)
    occupied: list[tuple[int, int]] = []
    matched: list[dict[str, str]] = []

    for term in ranked:
        needle = term["source_term"]
        if not needle:
            continue

        start = 0
        found = False

        while True:
            idx = text.find(needle, start)
            if idx < 0:
                break

            end = idx + len(needle)

            # Check for overlap with already matched longer terms
            if not any(idx < taken_end and end > taken_start for taken_start, taken_end in occupied):
                occupied.append((idx, end))
                matched.append(term)
                found = True
                break

            start = idx + 1

        if found:
            continue

    # Re-sort matched terms in chronological order of appearance in the string
    matched.sort(key=lambda item: text.find(item["source_term"]))
    return matched


def format_terms(terms: list[dict[str, str]]) -> str:
    """Formats matched glossary list into a clean readable string for LLM context."""
    if not terms:
        return ""
    return "; ".join(f"{item['source_term']}:{item['target_term']}" for item in terms)