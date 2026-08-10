"""Build Partnering-with-AI-PMI-2026.pptx from current deck + presenter-kit notes.

Branding matches Chromatic Guide HTML: Outfit, warm neutrals, terracotta accent,
square corners only (no rounded shapes).
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path.home() / ".cursor" / "skills" / "pptx-builder" / "scripts"))
import pptx_helpers as ph  # noqa: E402
from pptx_helpers import (  # noqa: E402
    Inches,
    PP_ALIGN,
    Pt,
    RGBColor,
    _rect,
    _set_text,
    _textbox,
    add_blank_slide,
    add_text_block,
    create_presentation,
)

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _p(tag: str) -> str:
    return f"{{{NS_P}}}{tag}"


class _AnimIds:
    def __init__(self) -> None:
        self._v = 0

    def next(self) -> str:
        self._v += 1
        return str(self._v)


def _sid(shape) -> int:
    return int(shape.shape_id)


# Auto-build pacing (After Previous). Click advances to the next slide.
BUILD_START_DELAY_MS = 300  # pause after slide lands
BUILD_STEP_DELAY_MS = 600  # pause between beats
BUILD_FADE_MS = 400  # fade duration


def _strip_timing(slide) -> None:
    slide_elem = slide._element
    for child in list(slide_elem):
        if child.tag == _p("timing"):
            slide_elem.remove(child)


def _timing_skeleton(slide_elem, ctr: _AnimIds):
    timing = etree.SubElement(slide_elem, _p("timing"))
    tn_lst = etree.SubElement(timing, _p("tnLst"))
    root_par = etree.SubElement(tn_lst, _p("par"))
    root_ctn = etree.SubElement(
        root_par, _p("cTn"), id=ctr.next(), dur="indefinite", restart="never", nodeType="tmRoot"
    )
    root_children = etree.SubElement(root_ctn, _p("childTnLst"))
    seq = etree.SubElement(root_children, _p("seq"), concurrent="1", nextAc="seek")
    seq_ctn = etree.SubElement(seq, _p("cTn"), id=ctr.next(), dur="indefinite", nodeType="mainSeq")
    seq_children = etree.SubElement(seq_ctn, _p("childTnLst"))
    prev = etree.SubElement(seq, _p("prevCondLst"))
    prev_cond = etree.SubElement(prev, _p("cond"), evt="onPrev", delay="0")
    etree.SubElement(etree.SubElement(prev_cond, _p("tgtEl")), _p("sldTgt"))
    nxt = etree.SubElement(seq, _p("nextCondLst"))
    next_cond = etree.SubElement(nxt, _p("cond"), evt="onNext", delay="0")
    etree.SubElement(etree.SubElement(next_cond, _p("tgtEl")), _p("sldTgt"))
    bld_lst = etree.SubElement(timing, _p("bldLst"))
    return seq_children, bld_lst


def _add_fade_in(
    mid_children,
    ctr: _AnimIds,
    spid: int,
    *,
    node_type: str,
    grp_id: int,
    para_idx: int | None = None,
) -> None:
    """Fade entrance (smoother for auto builds than hard Appear)."""
    effect_par = etree.SubElement(mid_children, _p("par"))
    effect_ctn = etree.SubElement(
        effect_par,
        _p("cTn"),
        id=ctr.next(),
        presetID="10",
        presetClass="entr",
        presetSubtype="0",
        fill="hold",
        grpId=str(grp_id),
        nodeType=node_type,
    )
    effect_st = etree.SubElement(effect_ctn, _p("stCondLst"))
    etree.SubElement(effect_st, _p("cond"), delay="0")
    effect_children = etree.SubElement(effect_ctn, _p("childTnLst"))

    p_set = etree.SubElement(effect_children, _p("set"))
    c_bhvr = etree.SubElement(p_set, _p("cBhvr"))
    set_ctn = etree.SubElement(c_bhvr, _p("cTn"), id=ctr.next(), dur="1", fill="hold")
    set_st = etree.SubElement(set_ctn, _p("stCondLst"))
    etree.SubElement(set_st, _p("cond"), delay="0")
    tgt = etree.SubElement(c_bhvr, _p("tgtEl"))
    sp_tgt = etree.SubElement(tgt, _p("spTgt"), spid=str(spid))
    if para_idx is not None:
        tx_el = etree.SubElement(sp_tgt, _p("txEl"))
        etree.SubElement(tx_el, _p("pRg"), st=str(para_idx), end=str(para_idx))
    attrs = etree.SubElement(c_bhvr, _p("attrNameLst"))
    etree.SubElement(attrs, _p("attrName")).text = "style.visibility"
    to_el = etree.SubElement(p_set, _p("to"))
    etree.SubElement(to_el, _p("strVal"), val="visible")

    anim = etree.SubElement(effect_children, _p("animEffect"), transition="in", filter="fade")
    a_bhvr = etree.SubElement(anim, _p("cBhvr"))
    etree.SubElement(a_bhvr, _p("cTn"), id=ctr.next(), dur=str(BUILD_FADE_MS))
    a_tgt = etree.SubElement(a_bhvr, _p("tgtEl"))
    a_sp = etree.SubElement(a_tgt, _p("spTgt"), spid=str(spid))
    if para_idx is not None:
        a_tx = etree.SubElement(a_sp, _p("txEl"))
        etree.SubElement(a_tx, _p("pRg"), st=str(para_idx), end=str(para_idx))


def _auto_step_group(
    seq_children,
    ctr: _AnimIds,
    shape_ids: list[int],
    *,
    is_first: bool,
    para_idxs: list[int | None] | None = None,
) -> None:
    """One auto beat: After Previous, shapes in the group With Previous."""
    outer_delay = str(BUILD_START_DELAY_MS if is_first else BUILD_STEP_DELAY_MS)
    outer_par = etree.SubElement(seq_children, _p("par"))
    outer_ctn = etree.SubElement(outer_par, _p("cTn"), id=ctr.next(), fill="hold")
    outer_st = etree.SubElement(outer_ctn, _p("stCondLst"))
    etree.SubElement(outer_st, _p("cond"), delay=outer_delay)
    outer_children = etree.SubElement(outer_ctn, _p("childTnLst"))
    mid_par = etree.SubElement(outer_children, _p("par"))
    mid_ctn = etree.SubElement(mid_par, _p("cTn"), id=ctr.next(), fill="hold")
    mid_st = etree.SubElement(mid_ctn, _p("stCondLst"))
    etree.SubElement(mid_st, _p("cond"), delay="0")
    mid_children = etree.SubElement(mid_ctn, _p("childTnLst"))

    for i, spid in enumerate(shape_ids):
        para_idx = None if para_idxs is None else para_idxs[i]
        node_type = "afterEffect" if i == 0 else "withEffect"
        _add_fade_in(
            mid_children,
            ctr,
            spid,
            node_type=node_type,
            grp_id=i,
            para_idx=para_idx,
        )


def add_click_appear_builds(slide, steps: list[list[int]]) -> None:
    """Auto-build: each step fades in After Previous (no click required)."""
    if not steps:
        return
    _strip_timing(slide)
    ctr = _AnimIds()
    seq_children, bld_lst = _timing_skeleton(slide._element, ctr)
    seen: set[int] = set()
    first = True
    for shape_ids in steps:
        if not shape_ids:
            continue
        _auto_step_group(seq_children, ctr, shape_ids, is_first=first)
        first = False
        for spid in shape_ids:
            if spid not in seen:
                etree.SubElement(bld_lst, _p("bldP"), spid=str(spid), grpId="0")
                seen.add(spid)


def add_bullet_appear_builds(slide, bullet_shape, trailing_steps: list[list[int]] | None = None) -> None:
    """Auto-build: each bullet fades in After Previous, then trailing groups."""
    paras = [
        i
        for i, p in enumerate(bullet_shape.text_frame.paragraphs)
        if (p.text or "").strip()
    ]
    if not paras and not trailing_steps:
        return

    _strip_timing(slide)
    ctr = _AnimIds()
    seq_children, bld_lst = _timing_skeleton(slide._element, ctr)
    spid = _sid(bullet_shape)
    first = True
    for para_idx in paras:
        _auto_step_group(
            seq_children,
            ctr,
            [spid],
            is_first=first,
            para_idxs=[para_idx],
        )
        first = False
    etree.SubElement(bld_lst, _p("bldP"), spid=str(spid), grpId="0", build="p")

    seen = {spid}
    for shape_ids in trailing_steps or []:
        if not shape_ids:
            continue
        _auto_step_group(seq_children, ctr, shape_ids, is_first=first)
        first = False
        for sid in shape_ids:
            if sid not in seen:
                etree.SubElement(bld_lst, _p("bldP"), spid=str(sid), grpId="0")
                seen.add(sid)


# Chromatic Guide / deck.css
FONT = "Outfit"
ph.DEFAULT_FONT = FONT
ph.TEXT_BLACK = RGBColor(0x1A, 0x1A, 0x1A)

CG = {
    "bg": RGBColor(0xFF, 0xFF, 0xFF),
    "text": RGBColor(0x1A, 0x1A, 0x1A),
    "text_body": RGBColor(0x33, 0x33, 0x33),
    "muted": RGBColor(0x88, 0x88, 0x88),
    "accent": RGBColor(0x7A, 0x3B, 0x2E),  # terracotta callout bar
    "callout_bg": RGBColor(0xF6, 0xF3, 0xED),
    "surface": RGBColor(0xEB, 0xE6, 0xDC),
    "panel": RGBColor(0xFA, 0xFA, 0xFA),
    "border": RGBColor(0xEC, 0xEC, 0xEC),
    "teal": RGBColor(0x4A, 0x66, 0x70),
    "ink": RGBColor(0x2A, 0x27, 0x24),
}

# Signature six-bar palette (title illustration / brand fingerprint)
BAR_COLORS = [
    RGBColor(0x5C, 0x4D, 0x3C),
    RGBColor(0x8B, 0x69, 0x14),
    RGBColor(0x6B, 0x53, 0x44),
    RGBColor(0x4A, 0x66, 0x70),
    RGBColor(0x5A, 0x4A, 0x6E),
    RGBColor(0x7A, 0x3B, 0x2E),
]

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
TOTAL_SLIDES = 12

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Partnering-with-AI-PMI-2026.pptx"
OUT_FALLBACK = ROOT / "Partnering-with-AI-PMI-2026-rebuild.pptx"
ATLAS_IMG = ROOT / "assets" / "atlas" / "atlas-dashboard.png"
IMG_DIR = ROOT / "images"
TITLE_IMG = IMG_DIR / "illust-title.png"
YOU_LEAD_IMG = IMG_DIR / "illust-you-lead.png"
NARRATIVE_IMG = IMG_DIR / "illust-narrative.png"


def add_picture(slide, path: Path, left, top, width=None, height=None) -> None:
    if not path.is_file():
        return
    kwargs = {"left": left, "top": top}
    if width is not None:
        kwargs["width"] = width
    if height is not None:
        kwargs["height"] = height
    slide.shapes.add_picture(str(path), **kwargs)


def content_slide(prs, title: str):
    """Master content slide: white bg + Outfit title (HTML-like light weight)."""
    slide = add_blank_slide(prs)
    tb = _textbox(slide, Inches(0.5), Inches(0.28), Inches(12.2), Inches(0.7))
    _set_text(tb, title, size=28, bold=False, color=CG["ink"])
    for p in tb.text_frame.paragraphs:
        p.font.name = FONT
    return slide


def chromatic_chrome(slide, page_num: int) -> None:
    """Master chrome: six-color bar footer + brand mark + page number (square segments)."""
    bar_h = Inches(0.11)
    bar_y = SLIDE_H - bar_h
    seg_w = SLIDE_W / len(BAR_COLORS)
    for i, c in enumerate(BAR_COLORS):
        _rect(slide, seg_w * i, bar_y, seg_w + Inches(0.02), bar_h, c, rounded=False)
    # Hairline above bars
    _rect(slide, Inches(0), bar_y - Inches(0.02), SLIDE_W, Inches(0.015), CG["border"], rounded=False)
    brand = _textbox(slide, Inches(0.45), Inches(7.05), Inches(3.5), Inches(0.28))
    tfb = _set_text(brand, "Chromatic Guide", size=10, bold=False, color=CG["muted"])
    for p in tfb.paragraphs:
        p.font.name = FONT
    tb = _textbox(slide, Inches(11.9), Inches(7.05), Inches(1.2), Inches(0.28))
    tf = _set_text(tb, f"{page_num} / {TOTAL_SLIDES}", size=10, bold=False, color=CG["muted"])
    for p in tf.paragraphs:
        p.font.name = FONT
        p.alignment = PP_ALIGN.RIGHT


def set_notes(slide, text: str) -> None:
    notes = slide.notes_slide
    tf = notes.notes_text_frame
    tf.clear()
    p = tf.paragraphs[0]
    first = True
    for block in text.strip().split("\n"):
        if first:
            p.text = block
            p.font.size = Pt(12)
            p.font.name = FONT
            first = False
        else:
            np = tf.add_paragraph()
            np.text = block
            np.font.size = Pt(12)
            np.font.name = FONT


def bullets(slide, items: list[str], left, top, width, height, size: int = 15):
    tb = _textbox(slide, left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.name = FONT
        p.font.size = Pt(size)
        p.font.color.rgb = CG["text_body"]
        p.space_before = Pt(8)
    return tb


def kicker(slide, text: str, top=Inches(1.0)) -> None:
    add_text_block(
        slide,
        Inches(0.5),
        top,
        Inches(12),
        Inches(0.32),
        [{"text": text.upper(), "size": 10, "bold": False, "color_key": "text_muted"}],
        colors,
    )


def branded_callout(slide, text: str, left, top, width, height):
    """HTML .deck-callout: warm fill + terracotta left bar, square corners."""
    bg = _rect(slide, left, top, width, height, CG["callout_bg"], rounded=False)
    bar = _rect(slide, left, top, Inches(0.09), height, CG["accent"], rounded=False)
    tb = _textbox(slide, left + Inches(0.32), top + Inches(0.2), width - Inches(0.5), height - Inches(0.32))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT
    p.font.size = Pt(16)
    p.font.bold = False
    p.font.color.rgb = CG["text"]
    return [_sid(bg), _sid(bar), _sid(tb)]


def square_panel(slide, left, top, width, height, fill):
    return _rect(slide, left, top, width, height, fill, rounded=False)


NOTES = [
    # 1
    """SAY
Some of you saw me last year on productivity prompts. That still works. Tonight is year two — and the story underneath it is ATLAS.
I lead the structure, the stories, and the standards. AI helps me get it done. And: Accelerate the work. Don't dilute the craft.

TRANSITION
Quick intro on what I do — Friday→Monday foreshadow.""",
    # 2
    """SAY
I lead the VMO for our Corporate Data and AI Office — manage, track, and report at a data center company.
I help run AI Superworkers — 1000+ people. Demos, surveys, lots of "here's where I got stuck."
We built ATLAS so Monday leadership status is live data — not a Friday-night PowerPoint rebuild. I'll show that later — it's the through-line tonight.
Last year was prompts that saved hours. This year I also pull stories across sources — and build when nothing fits. ATLAS is that story. I do this for a living. Tonight I'll share what's been useful.

TRANSITION
Sixty seconds on where AI is for PMs.""",
    # 3
    """SAY
Keep this short. AI can draft status, pull notes together, even help you build.
Easy to produce a lot of updates. That doesn't mean they're clear or accurate.
Deloitte: about 74% plan to use agents; only about 1 in 5 have solid governance. Stick to approved tools; read before you send.
We already lived this with PowerPoint and spreadsheets — more slides didn't mean more clarity. Same risk, faster. That's the craft problem that led us to ATLAS. (Save the craft tagline for the demo cut and close.)

TRANSITION
What I still do myself.""",
    # 4
    """SAY
Quality and accuracy — I read it before it goes out. Decisions — AI can suggest; I decide. People — hard conversations stay human.
Microsoft: 86% of AI users treat AI as a starting point — not the final answer. If nobody in the room can explain the update, we have a trust problem. That's why a Monday pulse only works if people can trust the numbers — which is why ATLAS has an audit trail and review built in.

TRANSITION
How I try things without making a mess — quick.""",
    # 5
    """SAY (60-75 sec max)
Hard cap. Try it on real work. Fix it or drop it. Pick the right tool. If I build — write the spec first, then review. That's how ATLAS got built.
Start on real work. Fix it or stop — don't send a draft you haven't read.

TRANSITION
Tools — forty-five seconds max.""",
    # 6
    """SAY (45 sec max)
I don't use one tool for everything. Copilot for daily work — watch for confident mistakes in email. Chat tools to think or pull a story. Cursor when I build — that's the ATLAS path: write what I need first, then review hard.
Use what your org allows. Same quality bar on every tool.

TRANSITION
Where AI helps me most.""",
    # 7
    """SAY
Too many notes, decks, threads. AI gives me a draft. I decide what leadership hears. Same job ATLAS does at portfolio scale — commitments, health, what's in the way.
Starting out: meeting notes → action table. Check every row before you send it. Here's how that showed up for our VMO — ATLAS.

TRANSITION
Into the capstone.""",
    # 8
    """SAY
This is the capstone of tonight's story — why we built it, how AI helped, how we run on it. Stay in the pain; don't product-tour.
What stopped: Friday-night PowerPoint rebuilds for Monday leadership.
ADO/Jira tracks the work. ATLAS answers what leadership asks — what we committed to, are we on track, what's in the way.
Build detail under ~90 seconds: specs first, tested as we went, Azure + SSO. Monday pulse is live data now. AI helped me build faster. The specs are why people trust it.

OPTIONAL NEAR-MISS (20–30s — your real story only)
One wrong or fluffy status line you almost shipped — what you caught and why. No invented KPIs.

TRANSITION
What that looks like on the screen.""",
    # 9
    """SAY
Point at the dashboard: This replaced Friday. Audit trail, live status, built a little at a time with review. Don't tour the product. Capstone proof — then the craft live.
Next — Cursor. Same review discipline that shipped ATLAS. Watch for the moment I cut something that's wrong.

ZOOM
Stop deck share → share Cursor.""",
    # 10
    """SAY
Not every Cursor button. Sources in a folder. Ask it to pull the story together. Cut what's wrong out loud — then keep going until I'd put my name on it. Same edit discipline that shipped ATLAS.
The craft moment isn't the prompt. It's the edit.

DEMO CLIMAX (MUST LAND)
When output arrives: find one wrong or fluffy line, delete it on screen, say why ("That's not true / that's not how we talk / leadership doesn't need this"). That cut is the whole thesis — and how ATLAS stayed trustworthy.
If live fails: open saved output and do the same cut.

TRANSITION
Re-share deck — a few things to try.""",
    # 11
    """SAY
ATLAS is my capstone for partnering with AI — not a product pitch for you to copy. Your Monday move is smaller and yours.
Before questions — pick one this-week action. Write it down or drop it in chat. Pause long enough for them to choose.
Four keepers + this week / this month panels. Close on the craft line only: Accelerate the work. Don't dilute the craft. (Identity already landed at open.)

TRANSITION
Resources in chat — then questions.""",
    # 12
    """SAY
Four take-homes — scan the QR or grab links from chat. No need to type URLs from the slide.
Close on the thesis: AI can accelerate the work. You still own the judgment, the story, and the standard.
I'll stay for questions.

ZOOM
Paste chat message now.""",
]


def main() -> None:
    global colors
    prs, colors = create_presentation(
        theme="light",
        use_template=False,
        custom_colors={
            "bg": CG["bg"],
            "primary": CG["ink"],
            "secondary": CG["teal"],
            "tertiary": CG["ink"],
            "accent": CG["accent"],
            "text_dark": CG["text"],
            "text_muted": CG["muted"],
            "surface": CG["surface"],
            "info_bg": CG["callout_bg"],
            "success_bg": CG["callout_bg"],
            "warning_bg": CG["callout_bg"],
            "divider": CG["border"],
        },
    )

    # --- 1 Title + agenda ---
    s = add_blank_slide(prs)
    add_text_block(
        s,
        Inches(0.5),
        Inches(0.7),
        Inches(7.8),
        Inches(0.9),
        [{"text": "Partnering with AI", "size": 38, "bold": True}],
        colors,
    )
    add_text_block(
        s,
        Inches(0.5),
        Inches(1.65),
        Inches(7.8),
        Inches(1.35),
        [
            {
                "text": "I lead the structure, the stories, and the standards. AI helps me get it done.",
                "size": 16,
            },
            {
                "text": "Accelerate the work. Don't dilute the craft.",
                "size": 16,
                "bold": True,
                "space_before": 8,
                "color_key": "accent",
            },
            {
                "text": "Melonie Poole · PMI Chapter · August 2026",
                "size": 13,
                "space_before": 12,
                "color_key": "text_muted",
            },
        ],
        colors,
    )
    add_text_block(
        s,
        Inches(0.5),
        Inches(3.3),
        Inches(7.8),
        Inches(0.4),
        [{"text": "What we'll cover", "size": 17, "bold": True}],
        colors,
    )
    bullets(
        s,
        [
            "A little about my role",
            "Where AI is right now (brief) — what it means for PMs",
            "How I use AI without lowering the bar",
            "ATLAS — how we manage, track, and report now",
            "Live Cursor + a few things to try this week",
        ],
        Inches(0.5),
        Inches(3.8),
        Inches(7.8),
        Inches(2.8),
        size=15,
    )
    add_picture(s, TITLE_IMG, Inches(8.6), Inches(1.6), width=Inches(4.3))
    chromatic_chrome(s, 1)
    set_notes(s, NOTES[0])

    # --- 2 Role ---
    s = content_slide(prs, "A little about my role")
    kicker(s, "Quick intro")
    bullets(
        s,
        [
            "I lead the VMO for our Corporate Data and AI Office — how we manage, track, and report on program work at a data center company",
            "I help run AI Superworkers, our company Community of Practice (1000+ people) — demos, surveys, and a lot of \"here's where I got stuck\"",
            "We built ATLAS with AI so Monday leadership status comes from live data — not a Friday-night PowerPoint rebuild. I'll show you that later",
        ],
        Inches(0.5),
        Inches(1.45),
        Inches(12.2),
        Inches(3.2),
    )
    branded_callout(
        s,
        "Last year was prompts that saved hours. This year I also pull stories across sources — and build when nothing fits. ATLAS is that story. I do this for a living. Tonight I'll share what's been useful.",
        Inches(0.5),
        Inches(5.0),
        Inches(12.2),
        Inches(1.5),
    )
    chromatic_chrome(s, 2)
    set_notes(s, NOTES[1])

    # --- 3 AI frame ---
    s = content_slide(prs, "Where AI is — and why PMs should care")
    kicker(s, "Quick frame")
    bullets(
        s,
        [
            "AI can do more than chat — draft status, pull notes together, even help you build something",
            "It's easy to produce a lot of updates and decks now — that doesn't mean they're clear or accurate",
            "Companies are racing ahead on agents — but only about 1 in 5 say they have solid governance for them",
        ],
        Inches(0.5),
        Inches(1.45),
        Inches(12.2),
        Inches(2.6),
    )
    branded_callout(
        s,
        "We already lived this with PowerPoint and spreadsheets — more slides didn't mean more clarity. Same risk, faster.",
        Inches(0.5),
        Inches(4.3),
        Inches(12.2),
        Inches(1.2),
    )
    add_text_block(
        s,
        Inches(0.5),
        Inches(5.7),
        Inches(12.2),
        Inches(0.6),
        [
            {
                "text": "Source: Deloitte State of AI in the Enterprise 2026 — ~74% plan agentic AI use; 21% report mature agent governance",
                "size": 11,
                "color_key": "secondary",
            }
        ],
        colors,
    )
    chromatic_chrome(s, 3)
    set_notes(s, NOTES[2])

    # --- 4 You still make the call (build) ---
    s = content_slide(prs, "You still make the call")
    kicker(s, "What I keep coming back to")
    add_picture(s, YOU_LEAD_IMG, Inches(0.35), Inches(1.25), width=Inches(6.4))
    bullet_shape = bullets(
        s,
        [
            "Quality & accuracy — I read it before it goes out",
            "Decisions — AI can suggest; I decide",
            "People — hard conversations stay human",
        ],
        Inches(7.0),
        Inches(1.5),
        Inches(5.7),
        Inches(2.5),
    )
    callout_ids = branded_callout(
        s,
        "86% of AI users say they treat AI as a starting point — not the final answer. If nobody in the room can explain the update, we have a trust problem.",
        Inches(0.5),
        Inches(4.7),
        Inches(12.2),
        Inches(1.2),
    )
    source_tb = add_text_block(
        s,
        Inches(0.5),
        Inches(6.1),
        Inches(12.2),
        Inches(0.4),
        [{"text": "Source: Microsoft Work Trend Index 2026", "size": 11, "color_key": "secondary"}],
        colors,
    )
    chromatic_chrome(s, 4)
    add_bullet_appear_builds(s, bullet_shape, [callout_ids, [_sid(source_tb)]])
    set_notes(s, NOTES[3])

    # --- 5 Working loop (build) ---
    s = content_slide(prs, "How I try things without making a mess")
    kicker(s, "How I work")
    steps = [
        ("1", "Try it on real work", "Something on my plate this week."),
        ("2", "Fix it or drop it", "If it's wrong, I fix it. If it isn't helping, I stop."),
        ("3", "Pick the right tool", "Copilot day-to-day · chat to think · Cursor when I build."),
        ("4", "If I build: write the spec first", "What it needs to do — then build — then review. That's how ATLAS got built."),
    ]
    y = 1.4
    build_steps: list[list[int]] = []
    for num, title, desc in steps:
        num_tb = add_text_block(
            s,
            Inches(0.5),
            Inches(y),
            Inches(0.5),
            Inches(0.55),
            [{"text": num, "size": 22, "bold": True}],
            colors,
        )
        body_tb = add_text_block(
            s,
            Inches(1.2),
            Inches(y),
            Inches(11.5),
            Inches(0.7),
            [
                {"text": title, "size": 16, "bold": True},
                {"text": desc, "size": 13, "space_before": 2},
            ],
            colors,
        )
        build_steps.append([_sid(num_tb), _sid(body_tb)])
        y += 0.85
    callout_ids = branded_callout(
        s,
        "Start on real work. Fix it or stop — don't send a draft you haven't read.",
        Inches(0.5),
        Inches(5.5),
        Inches(12.2),
        Inches(1.0),
    )
    build_steps.append(callout_ids)
    chromatic_chrome(s, 5)
    add_click_appear_builds(s, build_steps)
    set_notes(s, NOTES[4])

    # --- 6 Tools (build) ---
    s = content_slide(prs, "I don't use one tool for everything")
    cols = [
        ("Microsoft Copilot", "Notes, email, Word, Teams — daily work. Watch for confident-sounding mistakes in mail."),
        ("ChatGPT · Claude · Gemini · PMI Affinity", "Thinking something through or pulling a story across sources."),
        ("Cursor · Claude Code", "Building — write what you need first, then review hard."),
    ]
    x = 0.5
    build_steps = []
    for title, body in cols:
        panel = square_panel(s, Inches(x), Inches(1.4), Inches(3.95), Inches(3.3), CG["callout_bg"])
        text_tb = add_text_block(
            s,
            Inches(x + 0.2),
            Inches(1.6),
            Inches(3.55),
            Inches(2.9),
            [
                {"text": title, "size": 15, "bold": True},
                {"text": body, "size": 13, "space_before": 10},
            ],
            colors,
        )
        build_steps.append([_sid(panel), _sid(text_tb)])
        x += 4.2
    callout_ids = branded_callout(
        s,
        "Use what your org allows. Same quality bar on every tool.",
        Inches(0.5),
        Inches(5.2),
        Inches(12.2),
        Inches(1.0),
    )
    build_steps.append(callout_ids)
    chromatic_chrome(s, 6)
    add_click_appear_builds(s, build_steps)
    set_notes(s, NOTES[5])

    # --- 7 Narrative ---
    s = content_slide(prs, "Where AI helps me most: pulling the story together")
    add_text_block(
        s,
        Inches(0.5),
        Inches(1.15),
        Inches(6.2),
        Inches(0.7),
        [
            {
                "text": "Too many notes, decks, and threads. AI gives me a draft. I decide what leadership hears.",
                "size": 15,
            }
        ],
        colors,
    )
    bullets(
        s,
        [
            "Inputs — notes, decks, transcripts",
            "AI — a draft I can work with",
            "Me — what stays, what gets cut, what's accurate",
        ],
        Inches(0.5),
        Inches(2.0),
        Inches(6.2),
        Inches(2.0),
        size=15,
    )
    add_picture(s, NARRATIVE_IMG, Inches(6.8), Inches(1.3), width=Inches(6.0))
    branded_callout(
        s,
        "Starting out: meeting notes → action table. Check every row before you send it. Here's how that showed up for our VMO — ATLAS.",
        Inches(0.5),
        Inches(5.0),
        Inches(12.2),
        Inches(1.2),
    )
    chromatic_chrome(s, 7)
    set_notes(s, NOTES[6])

    # --- 8 ATLAS (build) ---
    s = content_slide(prs, "ATLAS")
    kicker(s, "How we manage, track, and report")
    add_text_block(
        s,
        Inches(0.5),
        Inches(1.4),
        Inches(12.2),
        Inches(0.45),
        [{"text": "For our VMO — commitments, health, and status in one place.", "size": 16}],
        colors,
    )
    bullet_shape = bullets(
        s,
        [
            "What stopped: Friday-night PowerPoint rebuilds for Monday leadership",
            "ADO / Jira tracks the work · ATLAS answers what leadership asks — what we committed to, are we on track, what's in the way",
            "We wrote the specs first, tested as we went, hosted on company Azure with SSO",
        ],
        Inches(0.5),
        Inches(2.0),
        Inches(12.2),
        Inches(2.6),
    )
    callout_ids = branded_callout(
        s,
        "Monday pulse is live data now. AI helped me build faster. The specs are why people trust it.",
        Inches(0.5),
        Inches(5.0),
        Inches(12.2),
        Inches(1.2),
    )
    chromatic_chrome(s, 8)
    add_bullet_appear_builds(s, bullet_shape, [callout_ids])
    set_notes(s, NOTES[7])

    # --- 9 Trust / screenshot (ATLAS breathes — large visual) ---
    s = content_slide(prs, "Built so people can trust the numbers")
    add_text_block(
        s,
        Inches(0.5),
        Inches(1.05),
        Inches(12.2),
        Inches(0.45),
        [
            {
                "text": "Audit trail · live status — not Friday-night slides · built a little at a time",
                "size": 14,
                "color_key": "text_muted",
            }
        ],
        colors,
    )
    if ATLAS_IMG.is_file():
        add_picture(s, ATLAS_IMG, Inches(0.7), Inches(1.55), width=Inches(11.9))
        add_text_block(
            s,
            Inches(0.7),
            Inches(5.45),
            Inches(11.9),
            Inches(0.3),
            [
                {
                    "text": "This replaced Friday — portfolio pulse and deliverables by month",
                    "size": 12,
                    "color_key": "accent",
                }
            ],
            colors,
        )
    else:
        square_panel(s, Inches(0.7), Inches(1.55), Inches(11.9), Inches(3.6), CG["callout_bg"])
        add_text_block(
            s,
            Inches(1.0),
            Inches(3.0),
            Inches(11),
            Inches(1.0),
            [{"text": "[Insert ATLAS dashboard screenshot]", "size": 16, "color_key": "text_muted"}],
            colors,
        )
    branded_callout(
        s,
        "Next — Cursor. Watch for the moment I cut something that's wrong.",
        Inches(0.5),
        Inches(5.9),
        Inches(12.2),
        Inches(0.85),
    )
    chromatic_chrome(s, 9)
    set_notes(s, NOTES[8])

    # --- 10 Cursor bridge (build) ---
    s = content_slide(prs, "How I actually work")
    kicker(s, "Live demo")
    add_text_block(
        s,
        Inches(0.5),
        Inches(1.4),
        Inches(12.2),
        Inches(0.5),
        [{"text": "Not every Cursor button — how I check the work before I use it.", "size": 16}],
        colors,
    )
    bullet_shape = bullets(
        s,
        [
            "Sources in a folder (or a short spec if I'm building)",
            "Ask it to pull the story together",
            "Cut what's wrong out loud — then keep going until I'd put my name on it",
        ],
        Inches(0.5),
        Inches(2.1),
        Inches(12.2),
        Inches(2.4),
    )
    callout_ids = branded_callout(
        s,
        "The craft moment isn't the prompt. It's the edit. Same discipline that shipped ATLAS.",
        Inches(0.5),
        Inches(5.0),
        Inches(12.2),
        Inches(1.2),
    )
    chromatic_chrome(s, 10)
    add_bullet_appear_builds(s, bullet_shape, [callout_ids])
    set_notes(s, NOTES[9])

    # --- 11 Close ---
    s = content_slide(prs, "A few things to try")
    add_text_block(
        s,
        Inches(0.5),
        Inches(1.15),
        Inches(12.2),
        Inches(0.45),
        [
            {
                "text": "ATLAS is my capstone for partnering with AI. Before questions — pick one this-week action. Write it down or drop it in chat.",
                "size": 15,
                "bold": True,
            }
        ],
        colors,
    )
    bullets(
        s,
        [
            "You still make the call — that doesn't go away",
            "If you build something, start with a spec",
            "Use more than one tool — pick what fits the job",
            "Don't skip accuracy or the people side",
        ],
        Inches(0.5),
        Inches(1.7),
        Inches(12.2),
        Inches(2.0),
        size=15,
    )
    square_panel(s, Inches(0.5), Inches(3.85), Inches(5.9), Inches(1.65), CG["callout_bg"])
    add_text_block(
        s,
        Inches(0.7),
        Inches(4.0),
        Inches(5.5),
        Inches(1.4),
        [
            {"text": "This week", "size": 16, "bold": True},
            {"text": "• One Copilot task — and you check the whole thing", "size": 13, "space_before": 6},
            {"text": "• One meeting → action table you actually send", "size": 13, "space_before": 4},
        ],
        colors,
    )
    square_panel(s, Inches(6.7), Inches(3.85), Inches(5.9), Inches(1.65), CG["callout_bg"])
    add_text_block(
        s,
        Inches(6.9),
        Inches(4.0),
        Inches(5.5),
        Inches(1.4),
        [
            {"text": "This month", "size": 16, "bold": True},
            {"text": "• Pull a story from a few different sources", "size": 13, "space_before": 6},
            {"text": "• Try Cursor or Claude Code — start with a written spec", "size": 13, "space_before": 4},
        ],
        colors,
    )
    branded_callout(
        s,
        "Accelerate the work. Don't dilute the craft.",
        Inches(0.5),
        Inches(5.7),
        Inches(12.2),
        Inches(0.9),
    )
    chromatic_chrome(s, 11)
    set_notes(s, NOTES[10])

    # --- 12 Resources ---
    s = content_slide(prs, "Resources & take-home")
    resources = [
        (
            "PM AI Prompt Library",
            "Ready-to-use prompts for meetings, status, risk reviews, and executive communication.",
        ),
        (
            "What Stays With the PM",
            "What AI can help with — and what remains your responsibility.",
        ),
        (
            "Before You Build, Write the Spec",
            "One-page template for defining requirements before using AI.",
        ),
        (
            "AI Experiment Playbook",
            "Simple steps for testing AI on real work and measuring value.",
        ),
    ]
    y = 1.15
    for title, desc in resources:
        add_text_block(
            s,
            Inches(0.5),
            Inches(y),
            Inches(9.4),
            Inches(0.35),
            [{"text": title, "size": 16, "bold": True}],
            colors,
        )
        add_text_block(
            s,
            Inches(0.5),
            Inches(y + 0.32),
            Inches(9.4),
            Inches(0.45),
            [{"text": desc, "size": 13, "color_key": "text_muted"}],
            colors,
        )
        y += 0.85
    add_text_block(
        s,
        Inches(0.5),
        Inches(4.7),
        Inches(9.4),
        Inches(0.35),
        [{"text": "Connect: LinkedIn · Chromatic Guide", "size": 14}],
        colors,
    )
    qr_path = ROOT.parent / "assets" / "resources-qr.png"
    if qr_path.is_file():
        add_picture(s, qr_path, Inches(10.35), Inches(1.3), width=Inches(2.4))
        add_text_block(
            s,
            Inches(10.35),
            Inches(3.85),
            Inches(2.4),
            Inches(0.35),
            [{"text": "Scan for all resources", "size": 11, "color_key": "text_muted"}],
            colors,
        )
    branded_callout(
        s,
        "AI can accelerate the work. You still own the judgment, the story, and the standard.",
        Inches(0.5),
        Inches(5.35),
        Inches(12.2),
        Inches(1.15),
    )
    chromatic_chrome(s, 12)
    set_notes(s, NOTES[11])

    try:
        prs.save(str(OUT))
        wrote = OUT
    except PermissionError:
        prs.save(str(OUT_FALLBACK))
        wrote = OUT_FALLBACK
        print(f"NOTE: {OUT.name} is open — wrote fallback instead.")
    print(f"Wrote {wrote}")
    print(f"Slides: {len(prs.slides)}")
    print(f"ATLAS image embedded: {ATLAS_IMG.is_file()}")
    print("Auto-build animations (Fade After Previous): slides 4, 5, 6, 8, 10")
    print("Slide 12: named take-homes + QR (no URLs on slide)")


if __name__ == "__main__":
    main()
