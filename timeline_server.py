#!/usr/bin/env python3
"""
Relationship Evidence Pack — PDF Server
Run:  python3 timeline_server.py
Then: open timeline_builder.html in your browser
"""
from http.server import HTTPServer, BaseHTTPRequestHandler
import json, io, base64
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether, PageBreak, Image as RLImage,
    AnchorFlowable
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

from PIL import Image as PILImage
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader

# ── PALETTE ───────────────────────────────────────────────────
DARK    = colors.HexColor("#1A1814")
GOLD    = colors.HexColor("#A07830")
GOLD_LT = colors.HexColor("#C8A96E")
GOLD_BG = colors.HexColor("#FDF8F0")
LT_BG   = colors.HexColor("#F7F5F0")
WHITE   = colors.white
BORDER  = colors.HexColor("#D8D0C8")
TEXT    = colors.HexColor("#2A2420")
MUTED   = colors.HexColor("#7A7268")
FAINT   = colors.HexColor("#B8B0A4")
GREEN   = colors.HexColor("#2E7D5E")
GREEN_BG= colors.HexColor("#EEF8F3")
RED     = colors.HexColor("#B94C39")
RED_BG  = colors.HexColor("#FDF5F3")
INFO_BG = colors.HexColor("#EEF4FA")
INFO    = colors.HexColor("#4A7FA5")

TYPE_META = {
    "flight":   {"label":"Flight / Travel",        "prefix":"FLT","accent":colors.HexColor("#4A7FA5"),"bg":colors.HexColor("#EEF4FA"),"icon":"✈"},
    "airbnb":   {"label":"Accommodation",           "prefix":"ACM","accent":colors.HexColor("#C5503A"),"bg":colors.HexColor("#FDF1EF"),"icon":"⌂"},
    "transfer": {"label":"Bank Transfer",           "prefix":"TRF","accent":colors.HexColor("#2E7D5E"),"bg":colors.HexColor("#EEF8F3"),"icon":"$"},
    "chat":     {"label":"Chat Highlight",          "prefix":"MSG","accent":colors.HexColor("#7B5EA7"),"bg":colors.HexColor("#F4F0FA"),"icon":"✉"},
    "photo":    {"label":"Photo / Memory",          "prefix":"PHT","accent":colors.HexColor("#B07828"),"bg":colors.HexColor("#FBF5EA"),"icon":"◉"},
    "video":    {"label":"Video Call",              "prefix":"VID","accent":colors.HexColor("#1E7B8A"),"bg":colors.HexColor("#EAF6F8"),"icon":"▶"},
    "gap":      {"label":"Gap Explanation",         "prefix":"GAP","accent":colors.HexColor("#9A9088"),"bg":colors.HexColor("#F4F2F0"),"icon":"…"},
}

REL_STATUS_LABELS = {
    "unmarried_partners":"Unmarried partners (2+ years)",
    "engaged":"Engaged / Fiancé(e)",
    "married":"Married",
    "civil_partnership":"Civil partnership",
}
COHABIT_LABELS = {
    "yes_current":"Currently living together",
    "yes_past":"Previously cohabited",
    "no_longdistance":"Long-distance relationship",
    "no_cultural":"Not cohabiting — cultural / family reasons",
}

def dmy_to_iso(d):
    """Convert DD/MM/YYYY to YYYY-MM-DD. Returns input unchanged if already ISO or unrecognised."""
    if not d: return d
    # Already ISO
    import re
    if re.match(r'^\d{4}-\d{2}-\d{2}$', d): return d
    m = re.match(r'^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})$', d)
    if m:
        day, mon, yr = m.group(1), m.group(2), m.group(3)
        return f"{yr}-{mon.zfill(2)}-{day.zfill(2)}"
    return d

def fmt_date(d):
    if not d: return "—"
    d = dmy_to_iso(d)
    try: return datetime.strptime(d,"%Y-%m-%d").strftime("%-d %B %Y")
    except: return d

def fmt_my(d):
    if not d: return ""
    d = dmy_to_iso(d)
    try: return datetime.strptime(d,"%Y-%m-%d").strftime("%B %Y")
    except: return d

# ── FILE HELPERS ──────────────────────────────────────────────
def b64_to_bytes(data_url):
    """Strip data URL prefix and return raw bytes."""
    if ',' in data_url:
        data_url = data_url.split(',', 1)[1]
    return base64.b64decode(data_url)

def fix_orientation(pil_img):
    """Rotate image according to EXIF orientation tag so it prints the right way up."""
    try:
        exif = pil_img._getexif()
        if exif:
            orientation = exif.get(274)  # 274 = Orientation tag
            rotations = {3: 180, 6: 270, 8: 90}
            if orientation in rotations:
                pil_img = pil_img.rotate(rotations[orientation], expand=True)
    except Exception:
        pass  # No EXIF data or unreadable — just use image as-is
    return pil_img

def image_thumbnail(b64_data_url, max_w_mm=55, max_h_mm=45):
    """Return a ReportLab Image flowable thumbnail from a base64 image."""
    try:
        raw = b64_to_bytes(b64_data_url)
        buf = io.BytesIO(raw)
        pil = PILImage.open(buf)
        pil = fix_orientation(pil)
        if pil.mode not in ('RGB', 'L'):
            pil = pil.convert('RGB')
        w_px, h_px = pil.size
        dpi = 96
        max_w_px = max_w_mm / 25.4 * dpi
        max_h_px = max_h_mm / 25.4 * dpi
        scale = min(max_w_px / w_px, max_h_px / h_px, 1.0)
        out_w = (w_px * scale / dpi) * 25.4 * mm
        out_h = (h_px * scale / dpi) * 25.4 * mm
        img_buf = io.BytesIO()
        pil.save(img_buf, format='JPEG', quality=85)
        img_buf.seek(0)
        return RLImage(img_buf, width=out_w, height=out_h)
    except Exception as e:
        print(f"[image_thumbnail error] {e}")
        return None

def image_full(b64_data_url, page_w_mm=160):
    """Return a full-width ReportLab Image flowable."""
    try:
        raw = b64_to_bytes(b64_data_url)
        buf = io.BytesIO(raw)
        pil = PILImage.open(buf)
        pil = fix_orientation(pil)
        if pil.mode not in ('RGB', 'L'):
            pil = pil.convert('RGB')
        w_px, h_px = pil.size
        dpi = 96
        max_w_px = page_w_mm / 25.4 * dpi
        scale = min(max_w_px / w_px, 1.0)
        out_w = (w_px * scale / dpi) * 25.4 * mm
        out_h = (h_px * scale / dpi) * 25.4 * mm
        img_buf = io.BytesIO()
        pil.save(img_buf, format='JPEG', quality=90)
        img_buf.seek(0)
        return RLImage(img_buf, width=out_w, height=out_h)
    except Exception as e:
        print(f"[image_full error] {e}")
        return None

def parse_page_range(spec, total_pages):
    """Parse '1-3', 'all', '1,3,5' into a list of 0-indexed page numbers."""
    spec = (spec or 'all').strip().lower()
    if spec == 'all':
        return list(range(total_pages))
    pages = set()
    for part in spec.split(','):
        part = part.strip()
        if '-' in part:
            a, b = part.split('-', 1)
            try:
                pages.update(range(int(a)-1, min(int(b), total_pages)))
            except: pass
        else:
            try:
                p = int(part) - 1
                if 0 <= p < total_pages: pages.add(p)
            except: pass
    return sorted(pages)

def pdf_page_to_image(pdf_bytes, page_index, dpi=120):
    """Convert a single PDF page to a PIL Image using pypdf + pillow rendering."""
    try:
        # Use pypdf to extract page as image via reportlab's pdf2image approach
        # Since we can't render PDF to image without poppler, we'll return None
        # and handle PDFs differently (embed directly)
        return None
    except:
        return None

# ── STYLES ────────────────────────────────────────────────────
def mk_style(name,**kw):
    defaults=dict(fontName="Helvetica",fontSize=9,textColor=TEXT,leading=13)
    defaults.update(kw)
    return ParagraphStyle(name,**defaults)

def S():
    return {
      "eyebrow":    mk_style("e", fontName="Helvetica",    fontSize=8,  textColor=GOLD,   charSpace=3, spaceAfter=4),
      "cover_h":    mk_style("ch",fontName="Times-Bold",   fontSize=22, textColor=WHITE,  leading=28),
      "cover_s":    mk_style("cs",fontName="Times-Italic", fontSize=11, textColor=colors.HexColor("#A09080"), leading=16),
      "pg_title":   mk_style("pt",fontName="Times-Bold",   fontSize=15, textColor=DARK,   leading=20, spaceBefore=0, spaceAfter=6),
      "sec_head":   mk_style("sh",fontName="Helvetica-Bold",fontSize=8, textColor=MUTED,  charSpace=2, spaceBefore=6, spaceAfter=8),
      "body":       mk_style("b"),
      "body_b":     mk_style("bb",fontName="Helvetica-Bold",fontSize=9, textColor=TEXT,   leading=13),
      "body_it":    mk_style("bi",fontName="Helvetica-Oblique",fontSize=9,textColor=MUTED,leading=13),
      "small":      mk_style("sm",fontSize=8,  textColor=MUTED, leading=11),
      "small_it":   mk_style("si",fontName="Helvetica-Oblique",fontSize=8,textColor=MUTED,leading=11),
      "th_w":       mk_style("tw",fontName="Helvetica-Bold",fontSize=8, textColor=WHITE,  leading=11),
      "th_wr":      mk_style("twr",fontName="Helvetica-Bold",fontSize=8,textColor=WHITE,  leading=11,alignment=TA_RIGHT),
      "ref_code":   mk_style("rc",fontName="Courier-Bold", fontSize=9,  textColor=WHITE,  leading=12),
      "entry_date": mk_style("ed",fontName="Times-Bold",   fontSize=12, textColor=TEXT,   leading=15),
      "entry_main": mk_style("em",fontName="Helvetica-Bold",fontSize=9.5,textColor=TEXT,  leading=13),
      "entry_det":  mk_style("edt",fontSize=8, textColor=MUTED, leading=11),
      "entry_note": mk_style("en",fontName="Helvetica-Oblique",fontSize=8.5,textColor=MUTED,leading=12),
      "year_lbl":   mk_style("yl",fontName="Times-Bold",   fontSize=16, textColor=WHITE,  leading=20),
      "stmt_q":     mk_style("sq",fontName="Helvetica-Bold",fontSize=8.5,textColor=GOLD,   leading=12, spaceBefore=6),
      "stmt_a":     mk_style("sa",fontName="Helvetica",    fontSize=9,  textColor=TEXT,   leading=13, leftIndent=8, spaceAfter=2),
      "disclaimer": mk_style("d", fontName="Helvetica-Oblique",fontSize=7.5,textColor=FAINT,leading=11,spaceBefore=10),
      "wk_ok":      mk_style("wo",fontName="Helvetica-Bold",fontSize=9, textColor=GREEN,  leading=13),
      "wk_warn":    mk_style("ww",fontName="Helvetica-Bold",fontSize=9, textColor=RED,    leading=13),
      "wk_info":    mk_style("wi",fontName="Helvetica-Bold",fontSize=9, textColor=INFO,   leading=13),
      "wk_desc":    mk_style("wd",fontSize=8.5,textColor=MUTED,leading=12),
      "idx_ref":    mk_style("ir",fontName="Courier-Bold", fontSize=9,  textColor=TEXT,   leading=12),
    }

# ── WEAKNESS CHECKS (mirror JS logic) ────────────────────────
def run_checks(entries, data):
    results = []
    non_gap = [e for e in entries if e.get("type") != "gap"]
    dated = [e for e in non_gap if e.get("date")]
    dated_sorted = sorted(dated, key=lambda e: dmy_to_iso(e.get("date","")))

    def add(label, desc, ok_msg, warn_msg, status):
        results.append({"label":label,"desc":desc,"ok_msg":ok_msg,"warn_msg":warn_msg,"status":status})

    # 1. Spans full relationship
    rel_start = dmy_to_iso(data.get("rel_start",""))
    status = "warn"
    if rel_start and dated_sorted:
        try:
            start_yr = int(dmy_to_iso(rel_start)[:4])
            first_yr = int(dmy_to_iso(dated_sorted[0]["date"])[:4])
            if first_yr <= start_yr + 1: status = "ok"
        except: pass
    add("Evidence spans full relationship",
        "Evidence should cover the entire relationship duration, not just recent months.",
        "✓  Earliest evidence is from the start of the relationship",
        "⚠  No early evidence found — add entries from the beginning of the relationship",
        status)

    # 2. Variety of types
    types_used = set(e.get("type") for e in non_gap)
    status = "ok" if len(types_used) >= 3 else "warn"
    add("Multiple evidence types used",
        "A mix of types (travel, calls, transfers, photos) is more convincing than one type alone.",
        f"✓  {len(types_used)} different evidence types used",
        f"⚠  Only {len(types_used)} evidence type(s) — aim for at least 3 different types",
        status)

    # 3. Volume
    n = len(non_gap)
    status = "ok" if n >= 6 else "warn"
    add("Sufficient volume of evidence",
        "Aim for at least 6–10 pieces of evidence. More is better.",
        f"✓  {n} entries submitted",
        f"⚠  Only {n} entries — aim for 6 or more, ideally 10+",
        status)

    # 4. Recent evidence
    from datetime import timedelta
    six_ago = (datetime.today() - timedelta(days=180)).strftime("%Y-%m-%d")
    recent = [e for e in dated if dmy_to_iso(e["date"]) >= six_ago]
    status = "ok" if len(recent) >= 2 else "warn"
    add("Recent evidence included (last 6 months)",
        "The relationship must be 'subsisting' at the time of application — recent contact evidence is critical.",
        f"✓  {len(recent)} entries within the last 6 months",
        f"⚠  Only {len(recent)} recent entries — add contact evidence from the last 6 months",
        status)

    # 5. Gaps explained
    gap_entries = [e for e in entries if e.get("type") == "gap"]
    s_gaps = data.get("s_gaps","")
    status = "ok" if (gap_entries or len(s_gaps) > 50) else "info"
    add("Evidence gaps explained",
        "Unexplained gaps in contact are a red flag. Address them proactively.",
        "✓  Gap explanation provided",
        "ℹ  No gap explanations added — if there are quiet periods, explain them",
        status)

    # 6. Financial connection
    transfers = [e for e in non_gap if e.get("type") == "transfer"]
    status = "ok" if len(transfers) >= 2 else "warn"
    add("Financial connection evidenced",
        "Shared financial activity strengthens the genuine relationship argument.",
        f"✓  {len(transfers)} bank transfer entries",
        f"⚠  {len(transfers)} transfer entries — add at least 2 with attached bank statement pages",
        status)

    # 7. Both statements
    s = data.get("s_how_met",""); a = data.get("a_how_met","")
    if len(s) > 100 and len(a) > 100: status = "ok"
    elif len(s) > 100 or len(a) > 100: status = "info"
    else: status = "warn"
    add("Both personal statements completed",
        "Both partners should write a personal statement giving context the documents cannot provide.",
        "✓  Both statements completed",
        "⚠  One or both statements are incomplete — this is important supporting evidence",
        status)

    # 8. Witness
    wit = data.get("wit_name","")
    status = "ok" if len(wit) > 2 else "info"
    add("Supporting witness provided",
        "Third-party corroboration from someone who knows you both as a couple.",
        f"✓  Witness: {wit}",
        "ℹ  No witness listed — a supporting letter from a friend or family member is recommended",
        status)

    # 9. Previous relationships declared
    sp = data.get("sponsor_prev","").lower(); ap = data.get("applicant_prev","").lower()
    if "divorce" in sp or "previous" in sp:
        status = "ok" if ("decree" in sp or "held" in sp) else "warn"
    else: status = "ok"
    add("Previous relationships declared",
        "Any previous marriage must be formally ended. Attach decree absolute if applicable.",
        "✓  Previous relationship status addressed",
        "⚠  Previous relationship mentioned but no decree absolute referenced — attach it",
        status)

    # 10. Cohabitation explained
    cohabited = data.get("cohabited","")
    if cohabited in ("yes_current","yes_past"): status = "ok"
    else:
        s_maintain = data.get("s_maintain","")
        status = "ok" if len(s_maintain) > 100 else "warn"
    add("Cohabitation or long-distance explained",
        "If you haven't lived together, prove the relationship is genuine and durable through other means.",
        "✓  Cohabitation or long-distance situation addressed",
        "⚠  Not cohabiting but no explanation in personal statement — explain this clearly",
        status)

    return results

# ── BUILD PDF ─────────────────────────────────────────────────
def build_pdf(data):
    buf = io.BytesIO()
    W_pg, H_pg = A4
    L, R, T, B = 20*mm, 20*mm, 16*mm, 16*mm
    W = W_pg - L - R

    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=L, rightMargin=R, topMargin=T, bottomMargin=B,
        title="Relationship Evidence Pack")

    st = S()
    story = []
    today = datetime.today().strftime("%-d %B %Y")

    sponsor   = data.get("sponsor","Sponsor")
    applicant = data.get("applicant","Applicant")
    rel_start = data.get("rel_start","")
    country   = data.get("country","")
    app_ref   = data.get("app_ref","")
    entries   = data.get("entries",[])
    non_gap   = [e for e in entries if e.get("type") != "gap"]

    def hr(before=6, after=8):
        return HRFlowable(width=W, thickness=1, color=BORDER, spaceAfter=after, spaceBefore=before)

    def dark_hr():
        return HRFlowable(width=W, thickness=2, color=GOLD, spaceAfter=0, spaceBefore=0)

    def section_header(text):
        tbl = Table([[Paragraph(text, mk_style("sh2",fontName="Helvetica-Bold",fontSize=8,textColor=WHITE,charSpace=2,leading=11))]],colWidths=[W])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),DARK),
            ("LEFTPADDING",(0,0),(-1,-1),14),
            ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
        ]))
        return tbl

    # ── PAGE 1: COVER ─────────────────────────────────────────
    cover = Table([
        [Paragraph("UNITED KINGDOM FAMILY VISA APPLICATION", st["eyebrow"])],
        [Paragraph(f"Relationship Evidence Pack<br/><i>{sponsor}</i> &amp; <i>{applicant}</i>", st["cover_h"])],
        [Paragraph(
            f"A complete evidence submission — personal statements, {len(non_gap)} pieces of evidence, "
            f"weakness self-check, and evidence index"
            + (f", from {fmt_my(rel_start)}" if rel_start else ""),
            st["cover_s"])],
    ], colWidths=[W])
    cover.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),DARK),
        ("LEFTPADDING",(0,0),(-1,-1),20),("RIGHTPADDING",(0,0),(-1,-1),20),
        ("TOPPADDING",(0,0),(0,0),20),("TOPPADDING",(0,1),(-1,-1),5),
        ("BOTTOMPADDING",(0,-1),(-1,-1),20),("BOTTOMPADDING",(0,0),(-1,-2),3),
    ]))
    story.append(cover)
    story.append(dark_hr())

    # Meta strip
    meta_parts = []
    if sponsor:   meta_parts.append(f"<b>Sponsor:</b> {sponsor}")
    if applicant: meta_parts.append(f"<b>Applicant:</b> {applicant}")
    if country:   meta_parts.append(f"<b>Country:</b> {country}")
    if rel_start: meta_parts.append(f"<b>Since:</b> {fmt_date(rel_start)}")
    rs = REL_STATUS_LABELS.get(data.get("rel_status",""),"")
    if rs: meta_parts.append(f"<b>Status:</b> {rs}")
    if app_ref:   meta_parts.append(f"<b>Ref:</b> {app_ref}")
    meta_parts.append(f"<b>Prepared:</b> {today}")
    meta_row = Table([[Paragraph("  ·  ".join(meta_parts),
        mk_style("meta",fontSize=8.5,textColor=MUTED,leading=13))]], colWidths=[W])
    meta_row.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F0EDE8")),
        ("LEFTPADDING",(0,0),(-1,-1),18),("RIGHTPADDING",(0,0),(-1,-1),18),
        ("TOPPADDING",(0,0),(-1,-1),9),("BOTTOMPADDING",(0,0),(-1,-1),9),
        ("BOX",(0,0),(-1,-1),0.5,BORDER),
    ]))
    story.append(meta_row)
    story.append(Spacer(1,14))

    # Document map
    story.append(Paragraph("DOCUMENT CONTENTS", st["sec_head"]))
    pages = [
        ("Page 1", "Cover & Evidence Summary", "This page — overview of application and evidence counts"),
        ("Page 2", "Sponsor's Personal Statement", f"Written by {sponsor}"),
        ("Page 3", "Applicant's Personal Statement", f"Written by {applicant}"),
        ("Page 4", "Weakness Self-Check", "10 known refusal risk areas — all addressed"),
        ("Page 5+", "Chronological Timeline", f"{len(non_gap)} pieces of evidence in date order"),
        ("Final pages", "Evidence Index & Attachment Checklist", "Reference codes cross-referenced to physical files"),
    ]
    pg_rows = [[Paragraph("<b>"+p+"</b>",mk_style("pr",fontName="Helvetica-Bold",fontSize=8,textColor=GOLD,leading=11)),
                Paragraph("<b>"+t+"</b>",mk_style("pt2",fontName="Helvetica-Bold",fontSize=8.5,textColor=TEXT,leading=11)),
                Paragraph(d,mk_style("pd",fontSize=8,textColor=MUTED,leading=11))] for p,t,d in pages]
    pg_tbl = Table(pg_rows, colWidths=[W*0.12, W*0.33, W*0.55])
    pg_tbl.setStyle(TableStyle([
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[LT_BG,WHITE]),
        ("GRID",(0,0),(-1,-1),0.5,BORDER),
        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
    ]))
    story.append(pg_tbl)
    story.append(Spacer(1,14))

    # Evidence summary
    story.append(Paragraph("EVIDENCE SUMMARY", st["sec_head"]))
    type_counts = {}
    for e in non_gap:
        t = e.get("type",""); type_counts[t]=type_counts.get(t,0)+1
    total_atts = sum(len(e.get("attachments",[])) for e in non_gap)

    sum_rows = [[Paragraph("<b>Evidence Type</b>",st["th_w"]),Paragraph("<b>Prefix</b>",st["th_w"]),
                 Paragraph("<b>Count</b>",mk_style("tc",fontName="Helvetica-Bold",fontSize=8,textColor=WHITE,alignment=TA_CENTER)),
                 Paragraph("<b>Date Range</b>",st["th_wr"])]]
    for tid,mi in TYPE_META.items():
        if tid=="gap" or tid not in type_counts: continue
        c=type_counts[tid]; acc=mi["accent"]
        typed=[e for e in non_gap if e.get("type")==tid and e.get("date")]
        dates=sorted((dmy_to_iso(e["date"]) for e in typed), key=lambda d: d)
        dr=f"{fmt_date(dates[0])} — {fmt_date(dates[-1])}" if len(dates)>=2 else (fmt_date(dates[0]) if dates else "—")
        atts=sum(len(e.get("attachments",[])) for e in non_gap if e.get("type")==tid)
        att_str=f"  ({atts} file{'s' if atts!=1 else ''})" if atts else ""
        sum_rows.append([
            Paragraph(f"{mi['icon']}  {mi['label']}",mk_style("si2",fontName="Helvetica-Bold",fontSize=8.5,textColor=acc,leading=12)),
            Paragraph(mi["prefix"],mk_style("sp",fontName="Courier-Bold",fontSize=8.5,textColor=acc,leading=12)),
            Paragraph(str(c)+att_str,mk_style("sc",fontName="Helvetica-Bold",fontSize=8.5,textColor=acc,leading=12,alignment=TA_CENTER)),
            Paragraph(dr,mk_style("sd",fontSize=8,textColor=MUTED,leading=12,alignment=TA_RIGHT)),
        ])
    sum_rows.append([
        Paragraph("<b>TOTAL</b>",mk_style("tot",fontName="Helvetica-Bold",fontSize=8.5,textColor=GOLD)),
        Paragraph("",st["small"]),
        Paragraph(f"<b>{len(non_gap)}</b>",mk_style("tn",fontName="Helvetica-Bold",fontSize=8.5,textColor=GOLD,alignment=TA_CENTER)),
        Paragraph(f"{total_atts} attachment{'s' if total_atts!=1 else ''} listed",mk_style("ta",fontSize=8,textColor=MUTED,alignment=TA_RIGHT)),
    ])
    sum_tbl = Table(sum_rows, colWidths=[W*0.42,W*0.13,W*0.15,W*0.30])
    sum_tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),DARK),
        ("ROWBACKGROUNDS",(0,1),(-1,-2),[LT_BG,WHITE]),
        ("BACKGROUND",(0,-1),(-1,-1),GOLD_BG),
        ("GRID",(0,0),(-1,-1),0.5,BORDER),
        ("LEFTPADDING",(0,0),(-1,-1),10),("RIGHTPADDING",(0,0),(-1,-1),10),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story.append(sum_tbl)

    # ── PAGE 2: SPONSOR STATEMENT ─────────────────────────────
    story.append(PageBreak())
    story.append(section_header("SPONSOR'S PERSONAL STATEMENT"))
    story.append(Spacer(1,6))

    sp_name = sponsor
    story.append(Paragraph(
        f"I, <b>{sp_name}</b>, make the following statement in support of the visa application of <b>{applicant}</b>. "
        f"The information below is true and accurate to the best of my knowledge.",
        mk_style("intro",fontSize=9,textColor=MUTED,leading=14,spaceAfter=10)))

    stmt_fields_s = [
        ("How we met", data.get("s_how_met","")),
        ("How the relationship developed", data.get("s_develop","")),
        ("How we maintain the relationship day-to-day", data.get("s_maintain","")),
        ("Our future plans together in the UK", data.get("s_future","")),
        ("Explanation of any gaps in evidence", data.get("s_gaps","")),
        ("Additional information", data.get("s_other","")),
    ]
    for q, a in stmt_fields_s:
        if not a.strip(): continue
        story.append(Paragraph(q.upper(), st["stmt_q"]))
        story.append(Paragraph(a, st["stmt_a"]))

    # Previous relationships
    sp_prev = data.get("sponsor_prev","")
    if sp_prev:
        story.append(Paragraph("PREVIOUS RELATIONSHIPS", st["stmt_q"]))
        story.append(Paragraph(sp_prev, st["stmt_a"]))

    story.append(Spacer(1,8))
    story.append(hr())
    decl_s = Table([
        [Paragraph(f"Statement of truth — I confirm the above information is accurate to the best of my knowledge.  Signed: ________________________  Date: __________________", mk_style("ds",fontSize=8.5,textColor=MUTED,leading=13))],
    ], colWidths=[W])
    decl_s.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.5,BORDER),("LEFTPADDING",(0,0),(-1,-1),12),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),("BACKGROUND",(0,0),(-1,-1),LT_BG)]))
    story.append(decl_s)

    # ── PAGE 3: APPLICANT STATEMENT ───────────────────────────
    story.append(PageBreak())
    story.append(section_header("APPLICANT'S PERSONAL STATEMENT"))
    story.append(Spacer(1,6))
    story.append(Paragraph(
        f"I, <b>{applicant}</b>, make the following statement in support of my application for a UK Family Visa. "
        f"The information below is true and accurate to the best of my knowledge.",
        mk_style("intro2",fontSize=9,textColor=MUTED,leading=14,spaceAfter=10)))

    stmt_fields_a = [
        ("How we met", data.get("a_how_met","")),
        ("How the relationship developed", data.get("a_develop","")),
        ("How we maintain the relationship day-to-day", data.get("a_maintain","")),
        ("Why I want to live permanently in the UK", data.get("a_future","")),
        ("Explanation of any gaps in evidence", data.get("a_gaps","")),
    ]
    for q, a in stmt_fields_a:
        if not a.strip(): continue
        story.append(Paragraph(q.upper(), st["stmt_q"]))
        story.append(Paragraph(a, st["stmt_a"]))

    ap_prev = data.get("applicant_prev","")
    if ap_prev:
        story.append(Paragraph("PREVIOUS RELATIONSHIPS", st["stmt_q"]))
        story.append(Paragraph(ap_prev, st["stmt_a"]))

    # Witness section on applicant page
    wit_name = data.get("wit_name","")
    if wit_name:
        story.append(Spacer(1,12))
        story.append(section_header("SUPPORTING WITNESS"))
        story.append(Spacer(1,6))
        wit_rows = [
            [Paragraph("<b>Name</b>",st["small"]), Paragraph(wit_name, st["body"])],
            [Paragraph("<b>Relationship to couple</b>",st["small"]), Paragraph(data.get("wit_rel",""), st["body"])],
        ]
        if data.get("wit_summary"):
            wit_rows.append([Paragraph("<b>Letter covers</b>",st["small"]), Paragraph(data.get("wit_summary",""), st["body"])])
        wit_tbl = Table(wit_rows, colWidths=[W*0.25, W*0.75])
        wit_tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),LT_BG),
            ("GRID",(0,0),(-1,-1),0.5,BORDER),
            ("LEFTPADDING",(0,0),(-1,-1),10),("RIGHTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
            ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
        ]))
        story.append(wit_tbl)
        story.append(Paragraph("The witness's signed letter should be attached separately as a referenced document.",
            mk_style("wn",fontName="Helvetica-Oblique",fontSize=8,textColor=MUTED,leading=12,spaceBefore=6)))

    story.append(Spacer(1,8))
    story.append(hr())
    decl_a = Table([[Paragraph(f"Statement of truth — I confirm the above information is accurate to the best of my knowledge.  Signed: ________________________  Date: __________________",mk_style("da",fontSize=8.5,textColor=MUTED,leading=13))]],colWidths=[W])
    decl_a.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.5,BORDER),("LEFTPADDING",(0,0),(-1,-1),12),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),("BACKGROUND",(0,0),(-1,-1),LT_BG)]))
    story.append(decl_a)

    # ── PAGE 4: WEAKNESS CHECK ────────────────────────────────
    story.append(PageBreak())
    story.append(section_header("CASEWORKER WEAKNESS SELF-CHECK"))
    story.append(Spacer(1,8))
    story.append(Paragraph(
        "The following 10 checks address the most common reasons genuine UK family visa applications are refused. "
        "This page is included to demonstrate that known risk areas have been proactively considered and addressed.",
        mk_style("wk_intro",fontSize=9,textColor=MUTED,leading=14,spaceAfter=10)))

    checks = run_checks(entries, data)
    ok_count = sum(1 for c in checks if c["status"]=="ok")
    warn_count = sum(1 for c in checks if c["status"]=="warn")

    # Summary bar
    bar_color = GREEN if warn_count==0 else RED
    bar_bg = GREEN_BG if warn_count==0 else RED_BG
    bar_text = f"{ok_count}/10 checks passed" + (f"  ·  {warn_count} area{'s' if warn_count>1 else ''} to review" if warn_count else "  ·  All checks passed")
    bar = Table([[Paragraph(f"<b>{bar_text}</b>",mk_style("bar",fontName="Helvetica-Bold",fontSize=10,textColor=bar_color,leading=14))]],colWidths=[W])
    bar.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),bar_bg),("BOX",(0,0),(-1,-1),1,bar_color),("LEFTPADDING",(0,0),(-1,-1),14),("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8)]))
    story.append(bar)
    story.append(Spacer(1,10))

    for i, c in enumerate(checks):
        is_ok   = c["status"]=="ok"
        is_warn = c["status"]=="warn"
        icon    = "✓" if is_ok else ("⚠" if is_warn else "ℹ")
        acc     = GREEN if is_ok else (RED if is_warn else INFO)
        bg      = GREEN_BG if is_ok else (RED_BG if is_warn else INFO_BG)
        msg     = c["ok_msg"] if is_ok else c["warn_msg"]

        inner = [
            [Paragraph(f"{icon}  {c['label']}  —  {msg}",mk_style(f"wl{i}",fontName="Helvetica-Bold",fontSize=8.5,textColor=acc,leading=12))],
            [Paragraph(c["desc"],mk_style(f"wd{i}",fontSize=7.5,textColor=MUTED,leading=11))],
        ]
        inner_tbl=Table(inner,colWidths=[W-24])
        inner_tbl.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),1)]))
        row_tbl=Table([[inner_tbl]],colWidths=[W])
        row_tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),bg),
            ("LEFTPADDING",(0,0),(-1,-1),10),("RIGHTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
            ("LINEBEFORE",(0,0),(0,-1),3,acc),
            ("BOX",(0,0),(-1,-1),0.5,BORDER),
        ]))
        story.append(KeepTogether(row_tbl))
        story.append(Spacer(1,3))

    # ── PRE-BUILD GALLERY ANCHOR MAP ──────────────────────────
    # Map refCode -> anchor name for the first image with that code
    # Used to hyperlink ref codes in timeline and index to gallery
    gallery_anchors = {}
    for entry in entries:
        code = entry.get("refCode","")
        for att in entry.get("attachments",[]):
            if att.get("isImage") and att.get("b64") and code not in gallery_anchors:
                gallery_anchors[code] = f"img_{att['id']}"

    # ── TIMELINE ──────────────────────────────────────────────
    story.append(PageBreak())
    story.append(section_header("CHRONOLOGICAL TIMELINE OF RELATIONSHIP"))
    story.append(Spacer(1,8))

    prev_year = None
    for entry in entries:
        etype = entry.get("type","")
        mi = TYPE_META.get(etype,{"label":etype,"accent":MUTED,"bg":LT_BG,"icon":"•","prefix":"???"})
        acc=mi["accent"]; bg=mi["bg"]
        code=entry.get("refCode","???")
        edate=entry.get("date","")
        atts=entry.get("attachments",[])

        year=edate[:4] if edate else None
        if year and year!=prev_year:
            prev_year=year
            yd=Table([[Paragraph(year,st["year_lbl"])]],colWidths=[W])
            yd.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),DARK),("LEFTPADDING",(0,0),(-1,-1),14),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
            story.append(Spacer(1,6))
            story.append(yd)
            story.append(Spacer(1,5))

        main=entry.get("main","") or ""
        if etype=="transfer":
            amt=entry.get("amount",""); main=f"£{amt}" if amt else "Transfer"
            details=[entry.get("direction","")] if entry.get("direction") else []
            if entry.get("ref"): details.append(f"Ref: {entry['ref']}")
        elif etype=="flight": details=[f"Booking ref: {entry['ref']}"] if entry.get("ref") else []
        elif etype=="airbnb": details=([f"{entry['nights']} nights"] if entry.get("nights") else [])+(([f"Ref: {entry['ref']}"] if entry.get("ref") else []))
        elif etype=="video":  details=[f"Duration: {entry['duration']}"] if entry.get("duration") else []
        elif etype=="gap":    main=entry.get("gap_period","Gap in evidence"); details=[]
        else: details=[]

        note=entry.get("note","") or entry.get("summary","") or entry.get("description","") or entry.get("gap_reason","") or ""

        ref_text = f'<link dest="{gallery_anchors[code]}">{code}</link>' if code in gallery_anchors else code
        ref_bg=Table([[Paragraph(ref_text,st["ref_code"])]],colWidths=[18*mm])
        ref_bg.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),acc),("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3)]))
        left_rows=[[ref_bg],[Spacer(1,4)],[Paragraph(fmt_date(edate),st["entry_date"])]]
        left_tbl=Table(left_rows,colWidths=[W*0.27])
        left_tbl.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))

        right_rows=[[Paragraph(f"{mi['icon']}  {mi['label'].upper()}",mk_style(f"et{code}",fontName="Helvetica-Bold",fontSize=7,textColor=acc,charSpace=1,leading=10))]]
        if main:right_rows+=[[Spacer(1,3)],[Paragraph(main,st["entry_main"])]]
        if details:right_rows.append([Paragraph("  ·  ".join(details),st["entry_det"])])
        if note:right_rows+=[[Spacer(1,3)],[Paragraph(note,st["entry_note"])]]
        if atts:
            att_str="  ".join([f"[{a.get('fmt','?')}] {a.get('desc','')}" for a in atts])
            right_rows+=[[Spacer(1,4)],[Paragraph(f"📎  {att_str}",mk_style(f"al{code}",fontSize=7.5,textColor=acc,leading=11))]]

            # Collect thumbnails: real images + PDF placeholders
            thumb_items = []
            for a in atts:
                if a.get('isImage') and a.get('b64'):
                    thumb_items.append(('img', a))
                elif a.get('isPDF') and a.get('b64'):
                    thumb_items.append(('pdf', a))

            if thumb_items:
                thumb_cells=[]
                right_w = W * 0.70 - 16
                n_thumbs = min(len(thumb_items), 3)
                thumb_w = right_w / n_thumbs
                thumb_w_mm = thumb_w / mm

                for kind, a in thumb_items[:3]:
                    if kind == 'img':
                        thumb = image_thumbnail(a['b64'], max_w_mm=thumb_w_mm-3, max_h_mm=32)
                    else:
                        # PDF placeholder box
                        try:
                            ph_w = (thumb_w_mm - 3) * mm
                            ph_h = 32 * mm
                            ph = Table([[Paragraph(f"📄 PDF\n{a.get('desc','')[:20]}",
                                mk_style(f"ph{a['id']}",fontSize=7,textColor=MUTED,leading=10,alignment=TA_CENTER))]],
                                colWidths=[ph_w])
                            ph.setStyle(TableStyle([
                                ("BOX",(0,0),(-1,-1),0.5,BORDER),
                                ("BACKGROUND",(0,0),(-1,-1),LT_BG),
                                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                ("TOPPADDING",(0,0),(-1,-1),10),
                                ("BOTTOMPADDING",(0,0),(-1,-1),10),
                            ]))
                            thumb = ph
                        except:
                            thumb = None
                    if thumb:
                        cap = Paragraph(a.get('desc','')[:30], mk_style(f"tc{code}{a['id']}",fontSize=7,textColor=MUTED,leading=9,alignment=TA_CENTER))
                        thumb_cells.append([thumb, cap])

                if thumb_cells:
                    n = len(thumb_cells)
                    col_w = right_w / n
                    thumb_row=Table([[cell[0] for cell in thumb_cells]],colWidths=[col_w]*n)
                    cap_row=Table([[cell[1] for cell in thumb_cells]],colWidths=[col_w]*n)
                    thumb_row.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"BOTTOM"),("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),2),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
                    cap_row.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),2)]))
                    right_rows+=[[Spacer(1,5)],[thumb_row],[cap_row]]

        right_tbl=Table(right_rows,colWidths=[W*0.70])
        right_tbl.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),1)]))
        row_tbl=Table([[left_tbl,right_tbl]],colWidths=[W*0.30,W*0.70])
        row_tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),bg),
            ("LEFTPADDING",(0,0),(0,-1),8),("LEFTPADDING",(1,0),(1,-1),8),
            ("RIGHTPADDING",(0,0),(-1,-1),8),
            ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LINEBEFORE",(0,0),(0,-1),3,acc),
            ("BOX",(0,0),(-1,-1),0.5,BORDER),
        ]))
        story.append(KeepTogether(row_tbl))
        story.append(Spacer(1,3))

    # ── IMAGE GALLERY ─────────────────────────────────────────
    all_images = []
    for entry in entries:
        for att in entry.get("attachments",[]):
            if att.get("isImage") and att.get("b64"):
                all_images.append({"att":att,"refCode":entry.get("refCode","???"),"entry_main":entry.get("main","")})

    if all_images:
        story.append(PageBreak())
        story.append(section_header("IMAGE GALLERY"))
        story.append(Spacer(1,8))
        story.append(Paragraph(
            "All uploaded images in reference code order. Each image is scaled to fit the page.",
            mk_style("gi",fontSize=9,textColor=MUTED,leading=13,spaceAfter=10)))

        # Show 2 images per row, capped at half page width each
        col_w = (W - 10*mm) / 2

        def make_gallery_img(b64, max_w_mm, max_h_mm):
            try:
                raw = b64_to_bytes(b64)
                buf = io.BytesIO(raw)
                pil = PILImage.open(buf)
                pil = fix_orientation(pil)
                if pil.mode not in ('RGB','L'):
                    pil = pil.convert('RGB')
                w_px, h_px = pil.size
                # Convert max dimensions from mm to pixels at 96dpi for scaling
                dpi = 96
                max_w_px = max_w_mm / 25.4 * dpi
                max_h_px = max_h_mm / 25.4 * dpi
                scale = min(max_w_px / w_px, max_h_px / h_px, 1.0)
                # Final dimensions in ReportLab points
                out_w = (w_px * scale / dpi) * 25.4 * mm
                out_h = (h_px * scale / dpi) * 25.4 * mm
                img_buf = io.BytesIO()
                pil.save(img_buf, format='JPEG', quality=88)
                img_buf.seek(0)
                return RLImage(img_buf, width=out_w, height=out_h)
            except Exception as e:
                print(f"[gallery img error] {e}")
                return None

        # Group into pairs, placing anchors before each pair
        for i in range(0, len(all_images), 2):
            pair = all_images[i:i+2]
            # Place anchor before this pair (keyed to first item)
            story.append(AnchorFlowable(f"img_{pair[0]['att']['id']}"))

            cells = []
            for item in pair:
                att = item["att"]; code = item["refCode"]
                lbl = Paragraph(
                    f"<b>{code}</b>  {att.get('desc','')}",
                    mk_style(f"gl{att['id']}",fontSize=8,textColor=GOLD,leading=11,spaceAfter=3))
                img = make_gallery_img(att["b64"], max_w_mm=(col_w/mm)-4, max_h_mm=70)
                cell_content = [lbl, img] if img else [lbl]
                cells.append(cell_content)
            while len(cells) < 2:
                cells.append([Spacer(1,1)])
            tbl = Table([[cells[0], cells[1]]], colWidths=[col_w, col_w])
            tbl.setStyle(TableStyle([
                ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
                ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),8),
            ]))
            story.append(tbl)
            story.append(Spacer(1,4))
    # ── EVIDENCE INDEX ────────────────────────────────────────
    story.append(PageBreak())
    story.append(section_header("EVIDENCE INDEX"))
    story.append(Spacer(1,6))
    story.append(Paragraph(
        "Every reference code is listed below with the corresponding file(s) to attach. "
        "Name each file using the reference code shown — e.g. FLT-001_booking-confirmation.pdf — and submit in a clearly labelled folder.",
        mk_style("ii",fontSize=9,textColor=MUTED,leading=13,spaceAfter=10)))

    idx_rows=[[Paragraph("<b>Ref</b>",st["th_w"]),Paragraph("<b>Date</b>",st["th_w"]),Paragraph("<b>Type</b>",st["th_w"]),Paragraph("<b>Description</b>",st["th_w"]),Paragraph("<b>File(s) to attach</b>",st["th_w"])]]
    for entry in entries:
        if entry.get("type")=="gap": continue
        mi=TYPE_META.get(entry.get("type",""),{"label":"?","accent":MUTED})
        acc=mi["accent"]; code=entry.get("refCode","???")
        atts=entry.get("attachments",[])
        main=entry.get("main","") or entry.get("note","") or entry.get("summary","") or "—"
        if atts:
            att_p=Paragraph("<br/>".join([f"[{a.get('fmt','?')}]  {code}_{a.get('desc','file').lower().replace(' ','-')}" for a in atts]),mk_style(f"af{code}",fontName="Courier",fontSize=7.5,textColor=TEXT,leading=12))
        else:
            att_p=Paragraph("— no files listed —",st["small_it"])
        ref_text = f'<link dest="{gallery_anchors[code]}">{code}</link>' if code in gallery_anchors else code
        idx_rows.append([
            Paragraph(ref_text,mk_style(f"ic{code}",fontName="Courier-Bold",fontSize=9,textColor=acc,leading=12)),
            Paragraph(fmt_date(entry.get("date","")),mk_style(f"id{code}",fontSize=8,textColor=MUTED,leading=11)),
            Paragraph(mi["label"],mk_style(f"il{code}",fontName="Helvetica-Bold",fontSize=8,textColor=acc,leading=11)),
            Paragraph((main[:80]+"…" if len(main)>80 else main),mk_style(f"im{code}",fontSize=8.5,textColor=TEXT,leading=12)),
            att_p,
        ])
    idx_tbl=Table(idx_rows,colWidths=[W*0.12,W*0.14,W*0.14,W*0.27,W*0.33])
    idx_tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),DARK),("ROWBACKGROUNDS",(0,1),(-1,-1),[LT_BG,WHITE]),("GRID",(0,0),(-1,-1),0.5,BORDER),("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(idx_tbl)

    # ── ATTACHMENT CHECKLIST ──────────────────────────────────
    story.append(PageBreak())
    story.append(section_header("ATTACHMENT CHECKLIST"))
    story.append(Spacer(1,6))
    story.append(Paragraph("Tick each box when the file has been gathered, renamed to match the reference code, and added to the submission folder.",mk_style("ci",fontSize=9,textColor=MUTED,leading=13,spaceAfter=12)))

    for tid,mi in TYPE_META.items():
        if tid=="gap": continue
        type_entries=[e for e in entries if e.get("type")==tid]
        if not type_entries: continue
        acc=mi["accent"]
        hdr=Table([[Paragraph(f"{mi['icon']}  {mi['label'].upper()}",mk_style(f"clh{tid}",fontName="Helvetica-Bold",fontSize=8,textColor=WHITE,charSpace=1,leading=11))]],colWidths=[W])
        hdr.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),acc),("LEFTPADDING",(0,0),(-1,-1),10),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
        story.append(hdr)
        cl_rows=[[Paragraph("<b>☐</b>",mk_style("cb",fontName="Helvetica-Bold",fontSize=10,textColor=acc)),Paragraph("<b>Ref</b>",mk_style("cr",fontName="Helvetica-Bold",fontSize=8,textColor=MUTED)),Paragraph("<b>Entry</b>",mk_style("ce",fontName="Helvetica-Bold",fontSize=8,textColor=MUTED)),Paragraph("<b>File to attach</b>",mk_style("cf",fontName="Helvetica-Bold",fontSize=8,textColor=MUTED)),Paragraph("<b>Format</b>",mk_style("cfmt",fontName="Helvetica-Bold",fontSize=8,textColor=MUTED))]]
        for entry in type_entries:
            code=entry.get("refCode","???"); main=entry.get("main","") or "—"; atts=entry.get("attachments",[])
            if atts:
                for att in atts:
                    fname=f"{code}_{att.get('desc','file').lower().replace(' ','-')}"
                    cl_rows.append([Paragraph("☐",mk_style(f"cb2",fontSize=11,textColor=FAINT)),Paragraph(code,mk_style(f"crc{code}",fontName="Courier-Bold",fontSize=8.5,textColor=acc)),Paragraph((main[:40]+"…" if len(main)>40 else main),mk_style(f"cm{code}",fontSize=8.5,textColor=TEXT,leading=12)),Paragraph(fname,mk_style(f"cfn{code}",fontName="Courier",fontSize=7.5,textColor=MUTED,leading=11)),Paragraph(att.get("fmt","?"),mk_style(f"cff{code}",fontName="Helvetica-Bold",fontSize=8,textColor=acc))])
            else:
                cl_rows.append([Paragraph("☐",mk_style(f"cb3",fontSize=11,textColor=FAINT)),Paragraph(code,mk_style(f"crc2{code}",fontName="Courier-Bold",fontSize=8.5,textColor=acc)),Paragraph((main[:40]+"…" if len(main)>40 else main),mk_style(f"cm2{code}",fontSize=8.5,textColor=TEXT,leading=12)),Paragraph("(no file listed)",st["small_it"]),Paragraph("—",st["small"])])
        cl_tbl=Table(cl_rows,colWidths=[W*0.05,W*0.11,W*0.28,W*0.42,W*0.14])
        cl_tbl.setStyle(TableStyle([("ROWBACKGROUNDS",(0,0),(-1,-1),[LT_BG,WHITE]),("GRID",(0,0),(-1,-1),0.4,BORDER),("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F0EDE8"))]))
        story.append(cl_tbl)
        story.append(Spacer(1,8))

    # ── FINAL DECLARATION ─────────────────────────────────────
    story.append(Spacer(1,6))
    story.append(hr())
    story.append(Paragraph("JOINT DECLARATION", st["sec_head"]))
    decl_text=(f"We, <b>{sponsor}</b> (sponsor) and <b>{applicant}</b> (applicant), declare that all information and evidence in this document is true, accurate, and complete to the best of our knowledge. Our relationship is genuine and subsisting, and we intend to live together permanently in the United Kingdom.")
    decl_tbl=Table([[Paragraph(decl_text,mk_style("jd",fontSize=9,textColor=TEXT,leading=14))],[Spacer(1,14)],[Table([[Table([[Paragraph("Sponsor signature:",st["small"])],[HRFlowable(width=55*mm,thickness=0.5,color=BORDER)]],colWidths=[62*mm]),Table([[Paragraph("Applicant signature:",st["small"])],[HRFlowable(width=55*mm,thickness=0.5,color=BORDER)]],colWidths=[62*mm]),Table([[Paragraph("Date:",st["small"])],[HRFlowable(width=30*mm,thickness=0.5,color=BORDER)]],colWidths=[38*mm])]],colWidths=[W*0.37,W*0.37,W*0.26])]],colWidths=[W])
    decl_tbl.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.5,BORDER),("BACKGROUND",(0,0),(-1,-1),LT_BG),("LEFTPADDING",(0,0),(-1,-1),14),("RIGHTPADDING",(0,0),(-1,-1),14),("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),14)]))
    story.append(decl_tbl)
    story.append(Paragraph(f"Document prepared {today} to support a UK Family Visa application under Appendix FM. Not legal advice. For complex cases, consult a registered immigration adviser.",st["disclaimer"]))

    doc.build(story)
    buf.seek(0)
    main_pdf_bytes = buf.read()

    # ── MERGE EMBEDDED PDFs ───────────────────────────────────
    # Collect all PDF attachments that have b64 data
    pdf_attachments = []
    for entry in entries:
        for att in entry.get("attachments", []):
            if att.get("isPDF") and att.get("b64"):
                pdf_attachments.append(att)

    if not pdf_attachments:
        return main_pdf_bytes

    # Merge: main PDF + selected pages from each attachment PDF
    writer = PdfWriter()

    # Add all pages from main PDF
    main_reader = PdfReader(io.BytesIO(main_pdf_bytes))
    for page in main_reader.pages:
        writer.add_page(page)

    # Add selected pages from each uploaded PDF
    for att in pdf_attachments:
        try:
            raw = b64_to_bytes(att["b64"])
            reader = PdfReader(io.BytesIO(raw))
            total = len(reader.pages)
            page_indices = parse_page_range(att.get("embedPages", "1-3"), total)
            for idx in page_indices:
                if 0 <= idx < total:
                    writer.add_page(reader.pages[idx])
        except Exception as e:
            pass  # Skip unreadable PDFs silently

    merged_buf = io.BytesIO()
    writer.write(merged_buf)
    merged_buf.seek(0)
    return merged_buf.read()

# ── HTTP SERVER ───────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    def log_message(self,fmt,*args): pass  # suppress default access log
    def do_OPTIONS(self):
        self.send_response(200);self._cors();self.end_headers()
    def do_GET(self):
        # Serve the builder HTML at /
        if self.path in ('/', '/timeline_builder.html'):
            try:
                here = os.path.dirname(os.path.abspath(__file__))
                with open(os.path.join(here, 'timeline_builder.html'), 'rb') as f:
                    html = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(html)))
                self.end_headers()
                self.wfile.write(html)
            except FileNotFoundError:
                self.send_error(404, "timeline_builder.html not found")
        else:
            self.send_error(404)
    def do_POST(self):
        if self.path!="/generate":self.send_error(404);return
        length = int(self.headers.get("Content-Length",0))
        body = self.rfile.read(length)
        print(f"[server] Received {len(body)/1024:.1f}KB")
        try:
            data = json.loads(body)
            entries = data.get("entries", [])
            img_count = sum(1 for e in entries for a in e.get("attachments",[]) if a.get("isImage") and a.get("b64"))
            print(f"[server] {len(entries)} entries, {img_count} images")
            pdf = build_pdf(data)
            print(f"[server] PDF generated: {len(pdf)/1024:.1f}KB")
            self.send_response(200);self._cors()
            self.send_header("Content-Type","application/pdf")
            self.send_header("Content-Disposition","attachment; filename=relationship-evidence-pack.pdf")
            self.send_header("Content-Length",str(len(pdf)))
            self.end_headers();self.wfile.write(pdf)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_response(500);self._cors()
            self.send_header("Content-Type","text/plain")
            self.end_headers();self.wfile.write(str(e).encode())
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin","*")
        self.send_header("Access-Control-Allow-Headers","Content-Type")
        self.send_header("Access-Control-Allow-Methods","POST, OPTIONS")

if __name__=="__main__":
    import os
    PORT = int(os.environ.get("PORT", 5678))
    server=HTTPServer(("0.0.0.0",PORT),Handler)
    print(f"\n✓  Evidence Pack PDF server running on port {PORT}")
    print(f"   Open http://localhost:{PORT} in your browser\n")
    try: server.serve_forever()
    except KeyboardInterrupt: print("\nServer stopped.")
