#!/usr/bin/env python3
"""Generates 100 variants of the ConfSecond slide page into versions/ plus index.html."""
import html, os, re, sys, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PPTX = "ConfSecond.pptx"
GH = "https://github.com/caozbm/TEST2/blob/claude/jolly-planck-xippiv/" + PPTX

def slide_text():
    with zipfile.ZipFile(os.path.join(ROOT, PPTX)) as z:
        xml = z.read("ppt/slides/slide1.xml").decode("utf8")
    paras = []
    for p in re.findall(r"<a:p>.*?</a:p>|<a:p .*?</a:p>", xml, flags=re.S):
        t = "".join(re.findall(r"<a:t>(.*?)</a:t>", p))
        if t.strip():
            paras.append(html.unescape(t).strip())
    return paras

# name, light bg, light fg, light accent, light surface, dark bg, dark fg, dark accent, dark surface
PAL = [
 ("Ink",     "#f6f7f9","#14181f","#1d4ed8","#e8ebf1","#0e1117","#e7eaf0","#7aa2ff","#1a1f2a"),
 ("Moss",    "#f3f6f1","#18221a","#2f6b3c","#e2eadf","#0f1611","#e3ebe1","#79c28a","#1a251d"),
 ("Plum",    "#f8f4f8","#241626","#8a2b8f","#ecdfed","#150d17","#efe3f0","#d58bd9","#241a27"),
 ("Rust",    "#faf4f0","#2a1810","#b2461c","#f0e0d6","#190f0a","#f2e4db","#f09263","#2a1b14"),
 ("Teal",    "#f0f7f7","#0f2224","#0e7c86","#d9ecec","#081617","#dcefef","#4fd0db","#12282a"),
 ("Slate",   "#f4f5f7","#1f2430","#4a5568","#e3e6ec","#12151c","#e4e7ee","#a3b0c8","#1e232d"),
 ("Gold",    "#fbf8ef","#2a2410","#9a7400","#f0e8c9","#17140a","#f1ecd6","#e6bf3a","#262110"),
 ("Rose",    "#fbf3f4","#2b1519","#c0264a","#f3dde1","#190b0e","#f4e1e5","#f07a95","#2a1519"),
 ("Indigo",  "#f3f3fb","#171a33","#4338ca","#e1e1f5","#0c0d1c","#e3e4f6","#9a94ff","#191b36"),
 ("Graphite","#f7f7f6","#1a1a19","#d9480f","#e8e8e5","#101010","#ebebe8","#ff8a4c","#1d1d1c"),
]
# display, body, google families query
FONTS = [
 ("Fraunces","Source Sans 3","Fraunces:wght@600;800&family=Source+Sans+3:wght@400;600"),
 ("Space Mono","DM Sans","Space+Mono:wght@700&family=DM+Sans:wght@400;600"),
 ("Playfair Display","Lato","Playfair+Display:wght@700;900&family=Lato:wght@400;700"),
 ("Archivo Black","Work Sans","Archivo+Black&family=Work+Sans:wght@400;600"),
 ("Bitter","IBM Plex Sans","Bitter:wght@700;800&family=IBM+Plex+Sans:wght@400;600"),
 ("Syne","Manrope","Syne:wght@700;800&family=Manrope:wght@400;600"),
 ("Libre Baskerville","Nunito Sans","Libre+Baskerville:wght@700&family=Nunito+Sans:wght@400;700"),
 ("Oswald","Karla","Oswald:wght@500;700&family=Karla:wght@400;700"),
 ("DM Serif Display","Outfit","DM+Serif+Display&family=Outfit:wght@400;600"),
 ("IBM Plex Mono","IBM Plex Sans","IBM+Plex+Mono:wght@600&family=IBM+Plex+Sans:wght@400;600"),
]
LAYOUTS = [
 ("stack",     "Single column"),
 ("split",     "Split panels"),
 ("slide",     "16:9 slide frame"),
 ("sidebar",   "Left rail"),
 ("banner",    "Top banner"),
 ("cards",     "Three cards"),
 ("centered",  "Centered poster"),
 ("terminal",  "Monospace sheet"),
 ("quote",     "Big statement"),
 ("timeline",  "Stepped list"),
]
CSS_LAYOUT = {
"stack":"main{max-width:42rem}.panel{display:grid;gap:1.25rem}",
"split":"main{max-width:60rem}.panel{display:grid;grid-template-columns:1fr 1fr;gap:1.5rem}.panel>*{min-width:0}.title{grid-column:1/-1}@media(max-width:600px){.panel{grid-template-columns:1fr}}",
"slide":"main{max-width:56rem}.panel{aspect-ratio:16/9;max-width:100%;background:var(--surface);border:1px solid var(--line);padding:clamp(1rem,4vw,3rem);display:flex;flex-direction:column;justify-content:center;gap:1rem}@media(max-width:600px){.panel{aspect-ratio:auto}}",
"sidebar":"main{max-width:58rem}.panel{display:grid;grid-template-columns:minmax(0,14rem) minmax(0,1fr);gap:2rem}.title{border-right:3px solid var(--accent);padding-right:1.25rem}@media(max-width:600px){.panel{grid-template-columns:1fr}.title{border-right:0;border-bottom:3px solid var(--accent);padding:0 0 1rem}}",
"banner":"main{max-width:56rem}.title{background:var(--accent);color:var(--bg);padding:1.5rem;margin-bottom:1.5rem}.title h1,.title p{color:var(--bg)}.panel{display:block}",
"cards":"main{max-width:62rem}.panel{display:grid;grid-template-columns:repeat(auto-fit,minmax(15rem,1fr));gap:1rem}.panel>*{min-width:0;background:var(--surface);padding:1.25rem;border-radius:.5rem}",
"centered":"main{max-width:40rem;text-align:center}.panel{display:grid;gap:1.25rem;justify-items:center}.fact{text-align:left}",
"terminal":"main{max-width:48rem}body{font-family:var(--mono)}.panel{background:var(--surface);border:1px solid var(--line);padding:1.25rem;display:grid;gap:1rem}.title h1::before{content:'$ ';color:var(--accent)}",
"quote":"main{max-width:50rem}.fact{font-size:clamp(1.25rem,3.4vw,1.9rem);line-height:1.35;border-left:6px solid var(--accent);padding-left:1.25rem}.panel{display:grid;gap:1.5rem}",
"timeline":"main{max-width:46rem}.panel{display:grid;gap:0;border-left:2px solid var(--accent);margin-left:.5rem}.panel>*{padding:0 0 1.5rem 1.5rem;position:relative}.panel>*::before{content:'';position:absolute;left:-.45rem;top:.4rem;width:.8rem;height:.8rem;border-radius:50%;background:var(--accent)}",
}

def page(n, paras, rel_link):
    li, pi = n % 10, (n // 10) % 10
    fi = (n * 3 + n // 10) % 10
    lname, ldesc = LAYOUTS[li]
    pn, lb, lf, la, ls, db, df, da, ds = PAL[pi]
    fd, fb, fq = FONTS[fi]
    head, sub, *rest = paras
    fact = " ".join(rest)
    e = html.escape
    m = re.search(r"\d{7,}", fact)
    big = ""
    if m:
        num = f"{int(m.group()):,}"
        big = f'<p class="stat"><span class="num">{num}</span><span class="unit">Cs-133 radiation periods = 1 second</span></p>'
    mono = "'IBM Plex Mono'" if "Mono" in fd else "ui-monospace,Menlo,monospace"
    dark = f"--bg:{db};--fg:{df};--accent:{da};--surface:{ds};--line:color-mix(in srgb,{df} 22%,transparent);color-scheme:dark"
    return f"""<title>ConfSecond Slide v{n:03d}</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family={fq}&display=swap">
<style>
/* Variant {n:03d}: layout "{ldesc}", palette {pn}, display {fd}, body {fb} */
:root{{--bg:{lb};--fg:{lf};--accent:{la};--surface:{ls};--line:color-mix(in srgb,{lf} 22%,transparent);
--display:'{fd}',Georgia,serif;--body:'{fb}',system-ui,sans-serif;--mono:{mono}}}
@media(prefers-color-scheme:dark){{:root:not([data-theme="light"]){{{dark}}}}}
:root[data-theme="dark"]{{{dark}}}
body{{background:var(--bg);color:var(--fg);font:16px/1.6 var(--body);padding-inline:16px;padding-block:2rem}}
main{{margin-inline:auto}}
h1{{font:800 clamp(1.8rem,6vw,3rem)/1.1 var(--display);margin:0;text-wrap:balance;letter-spacing:-.01em}}
.sub{{font:600 1.1rem var(--display);color:var(--accent);margin:.5rem 0 0;text-transform:uppercase;letter-spacing:.06em}}
.fact{{margin:0;max-width:65ch;overflow-wrap:anywhere}}
.stat{{margin:0;display:grid;gap:.15rem}}
.num{{font:800 clamp(1.5rem,5.5vw,2.6rem)/1.1 var(--display);color:var(--accent);font-variant-numeric:tabular-nums;overflow-wrap:anywhere}}
.unit{{font-size:.8rem;text-transform:uppercase;letter-spacing:.08em;opacity:.75}}
.meta{{display:flex;flex-wrap:wrap;gap:.5rem 1.25rem;margin-top:2rem;padding-top:1rem;border-top:1px solid var(--line);font-size:.85rem}}
a{{color:var(--accent);font-weight:600}}a:focus-visible{{outline:3px solid var(--accent);outline-offset:3px}}
.tag{{opacity:.7}}
{CSS_LAYOUT[lname]}
</style>
<main>
<section class="panel">
<header class="title"><h1>{e(head)}</h1><p class="sub">{e(sub.rstrip(', '))}</p></header>
<p class="fact">{e(fact)}</p>
{big}
</section>
<footer class="meta"><a href="{rel_link}">Open ConfSecond.pptx</a><span class="tag">Variant {n:03d} of 100 · {e(ldesc)} · {pn}</span><a href="{'../' if rel_link.startswith('../') else ''}index.html">All versions</a></footer>
</main>
"""

def main():
    paras = slide_text()
    os.makedirs(os.path.join(ROOT, "versions"), exist_ok=True)
    rows = []
    for n in range(1, 101):
        # standalone repo pages link to the repo file relatively
        body = page(n, paras, "../" + PPTX)
        doc = "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">" + body.replace("<main>", "</head><body><main>",1) + "</body></html>"
        open(os.path.join(ROOT, "versions", f"v{n:03d}.html"), "w", encoding="utf8").write(doc)
        rows.append(f'<li><a href="versions/v{n:03d}.html">v{n:03d}</a> <span>{LAYOUTS[n%10][1]} · {PAL[(n//10)%10][0]}</span></li>')
    idx = f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ConfSecond Versions</title>
<style>body{{font:16px/1.6 system-ui,sans-serif;max-width:48rem;margin:2rem auto;padding-inline:16px}}ul{{columns:2 14rem;list-style:none;padding:0}}li span{{opacity:.65;font-size:.85rem}}</style></head><body>
<h1>ConfSecond: 100 page versions</h1><p><a href="{PPTX}">Download {PPTX}</a></p><ul>{''.join(rows)}</ul></body></html>"""
    open(os.path.join(ROOT, "index.html"), "w", encoding="utf8").write(idx)

if __name__ == "__main__":
    main()
