"""
Gemini prompt template and robust response parser.

The prompt is built DYNAMICALLY from PLACEHOLDER_MAP so it always
includes every placeholder defined in the app.
"""

import re
import json

from utils.placeholder_map import PLACEHOLDER_MAP


# ============================================================
# Dynamic Gemini Prompt Builder
# ============================================================
def build_gemini_prompt(include_all: bool = True) -> str:
    """
    Build the Gemini prompt dynamically from PLACEHOLDER_MAP.

    Args:
        include_all: If True, include every placeholder. If False, only
                     include placeholders in the FIO + EWP + DOC groups.

    Returns:
        Formatted prompt string
    """
    # Source groups to include
    if include_all:
        keys_to_include = [p["key"] for p in PLACEHOLDER_MAP
                           if p["source"] in ("FIO", "EWP", "DOC")]
    else:
        keys_to_include = [p["key"] for p in PLACEHOLDER_MAP
                           if p["source"] in ("FIO", "EWP")]

    # Build the fields block
    fields_lines = []
    for p in PLACEHOLDER_MAP:
        if p["key"] not in keys_to_include:
            continue
        key = p["key"]
        label = p.get("label", "")
        source = p.get("source", "")
        # Show example if default present, otherwise empty
        default = p.get("default", "")
        example = default if default else ""
        # Compose description line
        desc = f"{label} ({source})"
        if example:
            fields_lines.append(f'  "{key}": "{example}"   // {desc}')
        else:
            fields_lines.append(f'  "{key}": ""   // {desc}')

    fields_block = "{\n" + ",\n".join(fields_lines) + "\n}"

    prompt = f"""CRITICAL OUTPUT INSTRUCTIONS:
- Return EXACTLY ONE JSON object
- Do NOT include any other JSON, examples, or schema echo
- Do NOT wrap in markdown code fences
- Do NOT include explanatory text before or after
- Use DOUBLE quotes for all keys and string values
- Do NOT include trailing commas
- For missing values, use null
- Output MUST start with `{{` and end with `}}` — nothing else

Example of correct output:
{{
  "OLT_SITE": "CDO_013_GPONA_02",
  "OLT_OM_VLAN": "734",
  "OLT_MGMT_IP": "10.168.196.226",
  "AG_MODEL": "CX600-X8(V8)"
}}

--- BEGIN EXTRACTION TASK ---

You are an expert network engineer reading a Facility Implementation Order (FIO) Excel file
and an Engineering Work Plan (EWP) topology image for a Nokia Lightspan MF-2 OLT integration.

TASK:
Extract ALL of the following {len(fields_lines)} data points and return them as a single
valid JSON object.

FIELDS TO EXTRACT:

{fields_block}

INSTRUCTIONS:
1. Extract each field from the FIO Excel and the EWP image.
2. If a value is not found, set it to null.
3. Return ONLY a valid JSON object — no markdown fences, no explanation.
4. For node names, use the exact format from the source (hyphens, underscores preserved).
5. For VSI names, use the exact pattern (e.g., OM-90000734, IPOE-193013368).
6. For IPs, use the exact format found.
7. For VLANs, use plain digits (no decimal places).

FILES ATTACHED:
- FIO: <paste FIO data or upload file>
- EWP: <upload EWP image>

OUTPUT: A single JSON object with all {len(fields_lines)} fields above.
"""
    return prompt


# ============================================================
# Static fallback prompt (in case dynamic build fails)
# ============================================================
GEMINI_PROMPT_TEMPLATE = build_gemini_prompt(include_all=True)


# ============================================================
# Robust JSON Parser (unchanged)
# ============================================================
def parse_ai_response(response_text: str) -> dict:
    """
    Robustly parse AI responses into a dict.

    Handles:
    - Pure JSON
    - JSON inside ```json ... ``` fences
    - Multiple JSON blocks (uses FIRST valid one)
    - Single quotes instead of double quotes
    - Trailing commas
    - Python literals: None, True, False
    - Leading/trailing explanatory text
    - Extra whitespace / newlines
    - Braces inside string literals

    Raises ValueError if no valid JSON object can be extracted.
    """
    if not response_text or not response_text.strip():
        raise ValueError("Empty response — nothing to parse")

    text = response_text.strip()

    # ------------------------------------------------------------
    # STEP 1: Strip markdown code fences if present
    # ------------------------------------------------------------
    fence_match = re.search(
        r"```(?:json|JSON|python|py)?\s*(.*?)```",
        text,
        re.DOTALL,
    )
    if fence_match:
        text = fence_match.group(1).strip()

    # ------------------------------------------------------------
    # STEP 2: Try strict JSON parse on the whole text
    # ------------------------------------------------------------
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return _clean_and_return(data)
    except json.JSONDecodeError:
        pass

    # ------------------------------------------------------------
    # STEP 3: Extract ALL balanced { ... } blocks and try each
    # ------------------------------------------------------------
    blocks = _find_all_json_blocks(text)

    if not blocks:
        raise ValueError(
            "No JSON object found in response. Expected a `{ ... }` block."
        )

    last_error = None

    # Try each block — first valid dict wins
    for block in blocks:
        # Attempt 3a: strict parse
        try:
            data = json.loads(block)
            if isinstance(data, dict):
                return _clean_and_return(data)
        except json.JSONDecodeError as e:
            last_error = e

        # Attempt 3b: apply fixes, then parse
        fixed = _fix_common_json_issues(block)
        try:
            data = json.loads(fixed)
            if isinstance(data, dict):
                return _clean_and_return(data)
        except json.JSONDecodeError as e:
            last_error = e

        # Attempt 3c: ast.literal_eval (Python dict)
        try:
            import ast
            data = ast.literal_eval(fixed)
            if isinstance(data, dict):
                return _clean_and_return(data)
        except Exception as e:
            last_error = e

    # If we get here, no block was valid
    snippet = blocks[0][:300].replace("\n", " ") if blocks else ""
    raise ValueError(
        f"Invalid JSON: {last_error}\n\n"
        f"Found {len(blocks)} block(s). First block snippet:\n{snippet}..."
    )


# ============================================================
# Balanced-Brace Scanner
# ============================================================
def _find_all_json_blocks(text: str) -> list:
    """
    Find ALL balanced { ... } blocks in the text using a brace counter.
    Respects strings so braces inside string literals don't break the count.
    """
    blocks = []
    depth = 0
    start_idx = None
    in_string = False
    string_char = None

    i = 0
    n = len(text)

    while i < n:
        ch = text[i]

        # Handle escape sequences inside strings
        if in_string and ch == "\\" and i + 1 < n:
            i += 2
            continue

        # Toggle string mode
        if ch in ('"', "'"):
            if not in_string:
                in_string = True
                string_char = ch
            elif string_char == ch:
                in_string = False
                string_char = None
            i += 1
            continue

        # Skip brace counting if inside a string
        if in_string:
            i += 1
            continue

        # Track brace depth
        if ch == "{":
            if depth == 0:
                start_idx = i
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0 and start_idx is not None:
                blocks.append(text[start_idx:i + 1])
                start_idx = None
            elif depth < 0:
                depth = 0
                start_idx = None

        i += 1

    return blocks


# ============================================================
# Common JSON Fixes
# ============================================================
def _fix_common_json_issues(text: str) -> str:
    """Apply heuristic fixes for common AI JSON mistakes."""
    if not text:
        return text

    # 1. Replace Python literals with JSON equivalents
    text = re.sub(r"\bNone\b", "null", text)
    text = re.sub(r"\bTrue\b", "true", text)
    text = re.sub(r"\bFalse\b", "false", text)

    # 2. Remove trailing commas before } or ]
    text = re.sub(r",(\s*[}\]])", r"\1", text)

    # 3. Convert single-quoted strings to double-quoted
    try:
        if "'" in text and '"' not in text:
            text = text.replace("'", '"')
        else:
            text = re.sub(r"'([^']*)'(\s*:)", r'"\1"\2', text)
            text = re.sub(r"(:\s*)'([^']*)'", r'\1"\2"', text)
            text = re.sub(r"(,\s*)'([^']*)'", r'\1"\2"', text)
    except Exception:
        pass

    # 4. Remove BOM / zero-width chars
    text = text.replace("\ufeff", "").replace("\u200b", "")

    # 5. Collapse multiple blank lines
    text = re.sub(r"\n\s*\n+", "\n", text)

    return text.strip()


# ============================================================
# Cleanup & Return
# ============================================================
def _clean_and_return(data) -> dict:
    """Ensure dict, remove None/empty, convert values to strings."""
    if not isinstance(data, dict):
        raise ValueError(
            f"Expected a JSON object (dict), got {type(data).__name__}"
        )

    data = {
        k: v for k, v in data.items()
        if v is not None and v != ""
    }

    cleaned = {}
    for k, v in data.items():
        if isinstance(v, (str, int, float, bool)):
            cleaned[k] = str(v)
        elif isinstance(v, list):
            cleaned[k] = ", ".join(str(x) for x in v)
        elif isinstance(v, dict):
            cleaned[k] = json.dumps(v)
        else:
            cleaned[k] = str(v)

    return cleaned
