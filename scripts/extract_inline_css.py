import hashlib
import os
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
TEMPLATES = ROOT / "templates"
STATIC = ROOT / "static"


STYLE_RE = re.compile(r"<style\b[^>]*>(?P<css>[\s\S]*?)</style>", re.IGNORECASE)
HEAD_CLOSE_RE = re.compile(r"</head\s*>", re.IGNORECASE)


def sha1(s: str) -> str:
    return hashlib.sha1(s.encode("utf-8")).hexdigest()


def is_template(path: Path) -> bool:
    try:
        path.relative_to(TEMPLATES)
        return True
    except Exception:
        return False


def css_filename_for(html_path: Path) -> str:
    stem = html_path.stem
    # Avoid clobbering existing shared CSS files
    return f"{stem}.page.css"


def link_tag(filename: str, template: bool) -> str:
    if template:
        return f'<link rel="stylesheet" href="{{{{ url_for(\'static\', filename=\'{filename}\') }}}}">'
    return f'<link rel="stylesheet" href="static/{filename}">'


def inject_link(html: str, tag: str) -> str:
    # Put link right before </head>
    m = HEAD_CLOSE_RE.search(html)
    if not m:
        # If there's no </head>, just prepend
        return tag + "\n" + html
    insert_at = m.start()
    # keep indentation minimal (existing files vary)
    return html[:insert_at] + tag + "\n" + html[insert_at:]


def write_file_if_changed(path: Path, content: str) -> bool:
    if path.exists():
        existing = path.read_text(encoding="utf-8")
        if sha1(existing) == sha1(content):
            return False
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    return True


def process_html(html_path: Path) -> dict | None:
    html = html_path.read_text(encoding="utf-8")
    matches = list(STYLE_RE.finditer(html))
    if not matches:
        return None

    # Concatenate all style blocks (preserves order)
    css_parts: list[str] = []
    for m in matches:
        css_parts.append(m.group("css").strip("\n"))
    css = "\n\n".join(css_parts).strip() + "\n"

    # Remove all <style> blocks from HTML
    new_html = STYLE_RE.sub("", html)

    filename = css_filename_for(html_path)
    css_path = STATIC / filename

    tag = link_tag(filename, is_template(html_path))
    # Avoid duplicate injection
    if filename not in new_html:
        new_html = inject_link(new_html, tag)

    html_changed = sha1(new_html) != sha1(html)
    css_written = write_file_if_changed(css_path, css)
    if html_changed:
        html_path.write_text(new_html, encoding="utf-8")

    return {
        "html": str(html_path.relative_to(ROOT)),
        "css": str(css_path.relative_to(ROOT)),
        "style_blocks": len(matches),
        "html_changed": html_changed,
        "css_written": css_written,
    }


def main() -> None:
    html_files = list(ROOT.rglob("*.html"))
    # Skip anything under static/
    html_files = [p for p in html_files if "static" not in p.parts]

    results = []
    for p in sorted(html_files):
        r = process_html(p)
        if r:
            results.append(r)

    print(f"Processed {len(results)} HTML files with inline <style>.")
    for r in results:
        print(f"- {r['html']} -> {r['css']} (blocks={r['style_blocks']}, html_changed={r['html_changed']}, css_written={r['css_written']})")


if __name__ == "__main__":
    main()

