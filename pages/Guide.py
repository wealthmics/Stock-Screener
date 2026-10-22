"""
pages/Guide.py  -  the platform guide inside the deep analysis app.

A new file: nothing existing is edited. Streamlit turns every file in a pages folder into
its own page, so this one opens at <app address>/Guide. It reads guide.md from the screener
website (written there each night by add_guide.py), so there is only one copy of the guide
to keep up to date.
"""
import re

import requests
import streamlit as st

GUIDE_URL = 'https://wealthmics.github.io/Stock-Screener/guide.md'
SCREENER_URL = 'https://wealthmics.github.io/Stock-Screener/'

st.set_page_config(page_title='Platform guide | MICS International', layout='wide')

st.markdown("""
<style>
.block-container{padding-top:1.4rem;max-width:1200px}
.gtop{background:#172b4d;color:#fff;border-bottom:3px solid #bf8f00;border-radius:8px;padding:16px 22px;margin-bottom:18px;
 display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap}
.gtop .org{font-size:12.5px;font-weight:600;color:#c9d3e3}
.gtop h1{margin:2px 0 0;font-size:22px;font-weight:650;color:#fff;padding:0}
.gtop a{color:#fff;font-weight:600;font-size:13px;text-decoration:none;border:1px solid #5d7299;border-radius:6px;padding:7px 12px;margin-left:6px}
.gtop a:hover{background:#24406e}
div[data-testid="stMarkdownContainer"] h2{color:#1f3864}
div[data-testid="stMarkdownContainer"] h3{color:#1f3864}
div[data-testid="stMarkdownContainer"] table{font-size:13.5px}
div[data-testid="stMarkdownContainer"] th{background:#1f3864;color:#fff}
</style>""", unsafe_allow_html=True)

st.markdown(
    '<div class="gtop"><div><div class="org">MICS International</div><h1>Platform guide</h1></div>'
    f'<div><a href="/" target="_self">Deep analysis</a><a href="{SCREENER_URL}" target="_blank">Screener</a>'
    f'<a href="{SCREENER_URL}guide.html" target="_blank">Printable guide</a></div></div>',
    unsafe_allow_html=True)


@st.cache_data(ttl=3600, show_spinner=False)
def load_guide():
    r = requests.get(GUIDE_URL, timeout=20)
    r.raise_for_status()
    r.encoding = 'utf-8'
    return r.text


def split_sections(md):
    """[(title, text)] for every ## section; the # title line is dropped."""
    md = re.sub(r'^# .*\n', '', md, count=1)
    parts = re.split(r'(?m)^## ', md)
    intro, out = parts[0].strip(), []
    for p in parts[1:]:
        title, _, body = p.partition('\n')
        out.append((title.strip(), body))
    return intro, out


def show(text):
    """Markdown, with each ```latex block drawn as a formula."""
    pieces = re.split(r'```latex\n(.*?)\n```', text, flags=re.S)
    for i, piece in enumerate(pieces):
        if i % 2:
            st.latex(piece)
        elif piece.strip():
            st.markdown(piece)


try:
    guide = load_guide()
except Exception as e:
    st.error('The guide could not be loaded from the screener website. It is written there '
             'by add_guide.py after the nightly build.')
    st.caption(f'Detail: {e}')
    st.stop()

intro, sections = split_sections(guide)
sections = [(t, b) for t, b in sections if t != 'Index']

c1, c2 = st.columns([2, 1])
titles = ['All sections'] + [t for t, _ in sections]
pick = c1.selectbox('Go to a section', titles, index=0)
q = c2.text_input('Search the guide', placeholder='for example beta, momentum, REIT')

shown = sections if pick == 'All sections' else [(t, b) for t, b in sections if t == pick]
if q.strip():
    needle = q.strip().lower()
    shown = [(t, b) for t, b in shown if needle in t.lower() or needle in b.lower()]
    st.caption(f'{len(shown)} section{"s" if len(shown) != 1 else ""} mention "{q.strip()}".')

if not shown:
    st.info('Nothing in the guide matches that search.')
if pick == 'All sections' and not q.strip():
    st.markdown(intro)
for title, body in shown:
    st.markdown(f'## {title}')
    show(body)
