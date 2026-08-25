#!/usr/bin/env python3
"""
build.py -- content/*.yaml 을 읽어 site/ 안에 완성된 홈페이지를 만든다.

사용법:
    python build.py            # CV(docx/PDF) + 홈페이지를 모두 새로 만든다
    python build.py --serve    # 만들고 나서 http://localhost:8000 으로 미리보기
    python build.py --no-cv    # 홈페이지만 (CV 는 건드리지 않음)

고칠 곳은 content/ 와 templates/ 뿐이다. site/ 는 매번 지워지고 다시 만들어진다.
"""
import argparse
import re
import shutil
import sys
from datetime import date
from pathlib import Path

import yaml
from jinja2 import Environment, FileSystemLoader, select_autoescape
from markupsafe import Markup, escape

ROOT = Path(__file__).resolve().parent
CONTENT = ROOT / "content"
TEMPLATES = ROOT / "templates"
ASSETS = ROOT / "assets"
OUT = ROOT / "site"

# 어떤 템플릿이 어떤 파일로 나가는지. profile.yaml 의 nav 와 짝을 맞춘다.
PAGES = [
    ("index.html",      "index.html",      "Home"),
    ("academics.html",  "academics.html",  "Academics"),
    ("activities.html", "activities.html", "Activities"),
    ("contact.html",    "contact.html",    "Contact Me"),
]

ME = "Yang, H."          # 저자 목록에서 굵게 표시할 이름


# --------------------------------------------------------------------------- 유틸
def load_yaml(name):
    with open(CONTENT / name, encoding="utf-8") as f:
        return yaml.safe_load(f)


def bold_me(authors: str) -> Markup:
    """저자 문자열 안의 'Yang, H.' 를 굵게 만든다. (HTML 로 안전하게 내보낸다)"""
    if not authors:
        return Markup("")
    safe = str(escape(authors))
    return Markup(safe.replace(ME, f"<strong>{ME}</strong>"))


def bold_year(text) -> Markup:
    """문자열 안의 마지막 네 자리 연도를 굵게 만든다. (2024. -> <strong>2024</strong>.)"""
    safe = str(escape(text))
    return Markup(re.sub(r"(\d{4})(?!.*\d{4})", r"<strong>\1</strong>", safe))


def md_to_html(text: str) -> str:
    """
    아주 작은 마크다운 변환기. 외부 라이브러리를 쓰지 않으려고 직접 만들었다.
    지원: 빈 줄로 나뉜 문단, **굵게**, *기울임*, [링크](주소), <!-- 주석 -->
    """
    text = re.sub(r"<!--.*?-->", "", text, flags=re.S)
    blocks = [b.strip() for b in re.split(r"\n\s*\n", text) if b.strip()]
    out = []
    for b in blocks:
        b = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2">\1</a>', b)
        b = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", b)
        b = re.sub(r"(?<!\*)\*([^*\n]+)\*(?!\*)", r"<em>\1</em>", b)
        b = b.replace("\n", " ")
        out.append(f"<p>{b}</p>")
    return Markup("\n".join(out))


def load_stories():
    """content/stories/*.md 를 읽어 {파일이름: HTML} 로 돌려준다."""
    stories = {}
    d = CONTENT / "stories"
    if d.is_dir():
        for f in sorted(d.glob("*.md")):
            key = f.stem.replace("-", "_")
            stories[key] = md_to_html(f.read_text(encoding="utf-8"))
    return stories


def check(profile, cv):
    """빠뜨리기 쉬운 것들을 미리 잡아 준다."""
    problems = []

    pdf = ASSETS / profile["cv_pdf"]
    if not pdf.exists():
        problems.append(f"CV 파일이 없습니다: assets/{profile['cv_pdf']}")

    # yaml 안에서 가리키는 이미지가 실제로 있는지
    def walk(node, path="cv"):
        if isinstance(node, dict):
            for k, v in node.items():
                if k in ("image", "portrait", "brand_logo", "intro_image", "booth_image") and isinstance(v, str):
                    if not (ASSETS / v).exists():
                        problems.append(f"이미지 없음: assets/{v}  ({path}.{k})")
                elif k == "images" and isinstance(v, list):
                    for i, im in enumerate(v):
                        if not (ASSETS / im).exists():
                            problems.append(f"이미지 없음: assets/{im}  ({path}.images[{i}])")
                else:
                    walk(v, f"{path}.{k}")
        elif isinstance(node, list):
            for i, v in enumerate(node):
                walk(v, f"{path}[{i}]")

    walk(cv, "cv")
    walk(profile, "profile")

    # 메뉴와 실제로 만들어지는 파일이 어긋나지 않는지
    made = {p[1] for p in PAGES}
    for item in profile.get("nav", []):
        if item["file"] not in made:
            problems.append(f"메뉴 '{item['title']}' 가 없는 파일을 가리킵니다: {item['file']}")

    return problems


# --------------------------------------------------------------------------- 빌드
def build():
    profile = load_yaml("profile.yaml")
    cv = load_yaml("cv.yaml")
    stories = load_stories()

    problems = check(profile, cv)
    if problems:
        print("확인이 필요합니다:")
        for p in problems:
            print("  -", p)
        print()

    env = Environment(
        loader=FileSystemLoader(str(TEMPLATES)),
        autoescape=select_autoescape(["html"]),
        trim_blocks=True,
        lstrip_blocks=False,
    )
    env.filters["bold_me"] = bold_me
    env.filters["bold_year"] = bold_year

    # 홈페이지를 굽기 전에 CV(docx/PDF)를 먼저 새로 만든다.
    # 그래야 DOWNLOAD FULL CV 버튼이 항상 최신 CV 를 가리킨다.
    if "--no-cv" not in sys.argv:
        import build_cv
        print("  CV 만드는 중...")
        try:
            build_cv.main()
        except Exception as e:
            print(f"  CV 생성 실패: {e}")
            print("  (홈페이지는 계속 만듭니다. 기존 CV 파일이 그대로 쓰입니다.)")

    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)

    # 자료 복사: assets/ 통째로 + style.css
    shutil.copytree(ASSETS, OUT / "assets", dirs_exist_ok=True)
    shutil.copy(TEMPLATES / "style.css", OUT / "assets" / "style.css")

    ctx = {
        "profile": profile,
        "cv": cv,
        "stories": stories,
        "built_on": date.today().isoformat(),
    }

    for template_name, out_name, title in PAGES:
        html = env.get_template(template_name).render(
            page_title=title, this_page=out_name, **ctx
        )
        (OUT / out_name).write_text(html, encoding="utf-8")
        print(f"  만듦  site/{out_name}")

    # GitHub Pages 가 Jekyll 로 다시 처리하지 않게 하는 표시
    (OUT / ".nojekyll").write_text("", encoding="utf-8")

    # 도메인을 정했다면 content/CNAME 에 한 줄 적어 두면 여기서 같이 나간다
    cname = CONTENT / "CNAME"
    if cname.exists():
        shutil.copy(cname, OUT / "CNAME")
        print(f"  만듦  site/CNAME  ({cname.read_text().strip()})")

    n_pub = len(cv.get("publications", []))
    n_conf = len(cv["conferences"]["oral"]) + len(cv["conferences"]["poster"])
    print(f"\n완료: 논문 {n_pub}편 / preprint {len(cv.get('preprints', []))}편 / 학회 {n_conf}건 / 저서 {len(cv.get('books', []))}권 "
          f"/ 특허 {len(cv['patents']['list'])}건 / 수상 {len(cv.get('awards', []))}건")
    return len(problems)


def serve():
    import http.server
    import socketserver
    import os
    os.chdir(OUT)
    with socketserver.TCPServer(("", 8000), http.server.SimpleHTTPRequestHandler) as httpd:
        print("\n미리보기: http://localhost:8000   (멈추려면 Ctrl+C)")
        httpd.serve_forever()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--serve", action="store_true", help="빌드 후 미리보기 서버를 켠다")
    ap.add_argument("--no-cv", action="store_true", help="CV(docx/PDF) 만들기를 건너뛴다")
    args = ap.parse_args()
    bad = build()
    if args.serve:
        serve()
    sys.exit(0)
