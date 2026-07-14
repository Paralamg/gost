import re
from typing import Optional


def format_value(val, float_format: Optional[str]) -> str:
    if isinstance(val, float):
        if float_format:
            return f"{val:{float_format}}"
        return str(val)
    return str(val)


def render_bold(paragraph, text: str) -> None:
    """Parse **bold** fragments into Word runs."""
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            paragraph.add_run(part)