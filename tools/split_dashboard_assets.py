from pathlib import Path
import re

root = Path(r"c:\excel-parser")
html_file = root / "templates" / "dashboard" / "index.html"
css_dir = root / "static" / "assets" / "css"
js_dir = root / "static" / "assets" / "js"
css_dir.mkdir(parents=True, exist_ok=True)
js_dir.mkdir(parents=True, exist_ok=True)

html = html_file.read_text(encoding="utf-8")
style_blocks = re.findall(r"<style[^>]*>(.*?)</style>", html, flags=re.IGNORECASE | re.DOTALL)
script_blocks = re.findall(r"<script[^>]*>(.*?)</script>", html, flags=re.IGNORECASE | re.DOTALL)

if style_blocks:
    css_content = "\n\n".join(style_blocks)
    (css_dir / "dashboard.css").write_text(css_content, encoding="utf-8")

if script_blocks:
    js_content = "\n\n".join(script_blocks)
    (js_dir / "dashboard.js").write_text(js_content, encoding="utf-8")

html = html.replace(
    "<link rel=\"stylesheet\" href=\"{{ url_for('static', filename='style/style.css') }}\">\n<link rel=\"stylesheet\" href=\"../../static/style/style.css\">",
    "<link rel=\"stylesheet\" href=\"{{ url_for('static', filename='assets/css/dashboard.css') }}\">",
)

html = re.sub(r"(?is)<style[^>]*>.*?</style>", "", html)
html = re.sub(r"(?is)<script[^>]*>.*?</script>", "", html)
html = html.replace(
    "</body>",
    "  <script src=\"{{ url_for('static', filename='assets/js/dashboard.js') }}\"></script>\n</body>",
)

html_file.write_text(html, encoding="utf-8")

print(f"Extraídos {len(style_blocks)} bloques CSS y {len(script_blocks)} bloques JS")
print(f"CSS: {css_dir / 'dashboard.css'}")
print(f"JS: {js_dir / 'dashboard.js'}")
