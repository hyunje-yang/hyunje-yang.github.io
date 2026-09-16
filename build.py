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
import hashlib
import json
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
#
# description 은 구글 검색 결과에 제목 아래 뜨는 두 줄이고, 링크를 채팅에 붙였을 때도
# 같이 뜬다. 비워 두면 구글이 본문에서 아무 문장이나 골라 쓰므로 직접 적어 둔다.
# 길이는 155자 안쪽이 좋다. 그보다 길면 뒤가 "..." 로 잘린다.
PAGES = [
    ("index.html", "index.html", "Home", 1.0,
     "Hyunje Yang is a PhD candidate at The University of Texas at Austin. "
     "He builds machine learning models for storm surge, compound flooding, "
     "and flood inundation mapping."),
    ("academics.html", "academics.html", "Academics", 0.9,
     "Publications, conference presentations, research experience, awards, and "
     "patents of Hyunje Yang, PhD candidate in civil engineering at UT Austin."),
    ("activities.html", "activities.html", "Activities", 0.6,
     "Music, volunteer work, and student leadership of Hyunje Yang, "
     "including the band Muguet."),
    ("contact.html", "contact.html", "Contact Me", 0.6,
     "How to reach Hyunje Yang at The University of Texas at Austin: email, "
     "ORCID, Google Scholar, and GitHub."),
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


# --------------------------------------------------------- 검색엔진에 알려 주는 것들
def base_url(profile) -> str:
    """https://hyunjeyang.com 처럼 끝에 슬래시가 없는 형태로 돌려준다."""
    return (profile.get("site_url") or "").rstrip("/")


def page_url(root: str, out_name: str) -> str:
    """index.html 은 주소에 파일 이름을 붙이지 않는다. 주소가 둘로 갈리기 때문이다."""
    if not root:
        return ""
    return f"{root}/" if out_name == "index.html" else f"{root}/{out_name}"


def make_jsonld(profile, cv, root: str) -> Markup:
    """
    구글이 읽는 '이 사람이 누구인가' 카드(schema.org Person)를 만든다.

    핵심은 sameAs 다. ORCID, Google Scholar, GitHub 주소를 한자리에 적어 두면
    구글이 흩어져 있는 기록을 같은 사람으로 묶는다. 이름으로 검색했을 때
    홈페이지가 위로 올라오는 데 가장 크게 작용하는 부분이다.
    """
    links = profile.get("links", {})
    same_as = [links[k]["url"] for k in ("orcid", "scholar", "github")
               if links.get(k, {}).get("url")]

    # 학력에서 모교를 뽑는다. 같은 학교가 학사·석사로 두 번 나오면 한 번만 넣는다.
    alumni, seen = [], set()
    for e in cv.get("education", []):
        school = e.get("school")
        if school and school != profile["affiliation"]["university"] and school not in seen:
            seen.add(school)
            alumni.append({"@type": "CollegeOrUniversity", "name": school})

    person = {
        "@context": "https://schema.org",
        "@type": "Person",
        "name": profile["name"],
        "givenName": "Hyunje",
        "familyName": "Yang",
        "jobTitle": profile["header_lines"][1],
        "description": profile.get("research_vision") or cv.get("research_vision", ""),
        "affiliation": {
            "@type": "CollegeOrUniversity",
            "name": profile["affiliation"]["university"],
            "url": "https://www.utexas.edu/",
        },
        "worksFor": {
            "@type": "ResearchOrganization",
            "name": profile["affiliation"]["lab"],
            "url": profile["affiliation"].get("lab_url", ""),
            "parentOrganization": {
                "@type": "CollegeOrUniversity",
                "name": profile["affiliation"]["university"],
            },
        },
        "knowsAbout": [i["title"] for i in cv.get("research_interests", []) if i.get("title")],
    }
    if root:
        person["url"] = f"{root}/"
        person["image"] = f"{root}/assets/img/og-card.jpg"
    if profile.get("email"):
        person["email"] = f"mailto:{profile['email']}"
    if alumni:
        person["alumniOf"] = alumni
    if same_as:
        person["sameAs"] = same_as
    if links.get("orcid", {}).get("url"):
        person["identifier"] = {
            "@type": "PropertyValue",
            "propertyID": "ORCID",
            "value": links["orcid"]["url"],
        }

    # HTML 안에 넣으므로 </script> 같은 조각이 태그로 읽히지 않게 막는다.
    text = json.dumps(person, ensure_ascii=False, indent=2)
    text = text.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    return Markup(text)


def write_robots_and_sitemap(root: str, today: str):
    """
    robots.txt  -- 크롤러에게 '다 봐도 된다' 고 알리고 지도의 위치를 가리킨다.
    sitemap.xml -- 페이지 목록. 링크를 타고 다니지 않아도 전부 찾을 수 있게 한다.

    site_url 이 비어 있으면 둘 다 만들지 않는다. 주소가 없으면 쓸모가 없기 때문이다.
    """
    if not root:
        return []

    robots = (
        "User-agent: *\n"
        "Allow: /\n"
        "\n"
        f"Sitemap: {root}/sitemap.xml\n"
    )
    (OUT / "robots.txt").write_text(robots, encoding="utf-8")

    rows = []
    for _, out_name, _, priority, _ in PAGES:
        rows.append(
            "  <url>\n"
            f"    <loc>{page_url(root, out_name)}</loc>\n"
            f"    <lastmod>{today}</lastmod>\n"
            f"    <priority>{priority}</priority>\n"
            "  </url>"
        )
    sitemap = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        + "\n".join(rows)
        + "\n</urlset>\n"
    )
    (OUT / "sitemap.xml").write_text(sitemap, encoding="utf-8")
    return ["robots.txt", "sitemap.xml"]


def check(profile, cv):
    """빠뜨리기 쉬운 것들을 미리 잡아 준다."""
    problems = []

    for key in ("cv_pdf", "cv_pdf_short"):
        if profile.get(key) and not (ASSETS / profile[key]).exists():
            problems.append(f"CV 파일이 없습니다: assets/{profile[key]}")

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

    # 링크 미리보기 카드. Times New Roman 이 있는 로컬에서만 새로 만들어지고,
    # GitHub Actions(리눅스)에서는 저장소에 올려 둔 jpg 를 그대로 쓴다.
    try:
        import build_og
        build_og.main()
    except Exception as e:
        print(f"  미리보기 카드 생성 실패: {e}  (기존 파일을 그대로 씁니다)")

    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)

    # 자료 복사: assets/ 통째로 + style.css
    shutil.copytree(ASSETS, OUT / "assets", dirs_exist_ok=True)
    shutil.copy(TEMPLATES / "style.css", OUT / "assets" / "style.css")

    # style.css 가 바뀌면 주소도 바뀌게 한다.
    # 안 그러면 브라우저가 예전 style.css 를 캐시에서 꺼내 써서
    # 새 HTML 에 옛 디자인이 입혀진 화면이 나온다.
    css_ver = hashlib.sha1((TEMPLATES / "style.css").read_bytes()).hexdigest()[:8]
    # 탭 아이콘(favicon)도 같은 이유로 주소에 버전을 붙인다.
    fav_ver = hashlib.sha1((ASSETS / "img" / "favicon-32.png").read_bytes()).hexdigest()[:8]

    ctx = {
        "profile": profile,
        "cv": cv,
        "stories": stories,
        "built_on": date.today().isoformat(),
        "css_ver": css_ver,
        "fav_ver": fav_ver,
    }

    root = base_url(profile)
    if not root:
        print("  건너뜀  canonical / sitemap (profile.yaml 의 site_url 이 비어 있습니다)")
    og_card = ASSETS / "img" / "og-card.jpg"
    og_image = f"{root}/assets/img/og-card.jpg" if root and og_card.exists() else ""
    jsonld = make_jsonld(profile, cv, root) if root else ""
    uni = profile["affiliation"]["university"]

    for template_name, out_name, title, _priority, desc in PAGES:
        home = out_name == "index.html"
        html = env.get_template(template_name).render(
            page_title=title,
            this_page=out_name,
            page_description=desc,
            canonical_url=page_url(root, out_name),
            og_type="profile" if home else "website",
            # 홈은 제목 자리에 "누구인지" 가 통째로 보이는 편이 낫다
            og_title=(f"{profile['name']} | {profile['header_lines'][1]}, {uni}"
                      if home else f"{title} | {profile['name']}"),
            og_image=og_image,
            jsonld=jsonld if home else "",   # 사람 정보는 홈에 한 번만 넣는다
            **ctx,
        )
        (OUT / out_name).write_text(html, encoding="utf-8")
        print(f"  만듦  site/{out_name}")

    for name in write_robots_and_sitemap(root, ctx["built_on"]):
        print(f"  만듦  site/{name}")

    # GitHub Pages 가 Jekyll 로 다시 처리하지 않게 하는 표시
    (OUT / ".nojekyll").write_text("", encoding="utf-8")

    # 도메인을 정했다면 content/CNAME 에 한 줄 적어 두면 여기서 같이 나간다
    cname = CONTENT / "CNAME"
    if cname.exists():
        shutil.copy(cname, OUT / "CNAME")
        print(f"  만듦  site/CNAME  ({cname.read_text().strip()})")

    # 검색엔진 소유 확인 파일 (Google Search Console, Bing 등).
    # content/verify/ 에 받은 파일을 그대로 넣어 두면 사이트 맨 위로 복사된다.
    # 예: content/verify/google1a2b3c.html -> https://hyunjeyang.com/google1a2b3c.html
    verify = CONTENT / "verify"
    if verify.is_dir():
        for f in sorted(verify.iterdir()):
            if f.is_file():
                shutil.copy(f, OUT / f.name)
                print(f"  만듦  site/{f.name}  (소유 확인 파일)")

    n_pub = len(cv.get("publications", []))
    n_conf = len(cv["conferences"]["oral"]) + len(cv["conferences"]["poster"])
    print(f"\n완료: 논문 {n_pub}편 / 심사 중 {len(cv.get('under_review', []))}편 / 준비 중 {len(cv.get('in_preparation', []))}편 / 학회 {n_conf}건 / 저서 {len(cv.get('books', []))}권 "
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
