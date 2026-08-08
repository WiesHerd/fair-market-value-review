#!/usr/bin/env python3
"""
HTML Email / Web Report Builder
Generates a single, self-contained HTML file (exhibit chart embedded as a base64
data URI, all CSS inline) suitable for pasting into an email body or opening
directly in a browser -- plus a matching PDF with real 0.6in page margins.

Reuses the same data-loading and exhibit-chart logic as build_adjustment_report.py
so the numbers in the email always match the DOCX report.

USAGE:
  python3 build_html_email.py \
    --name "Pediatric Critical Care" \
    --providers-file cohort.tsv \
    --market-anchors-file anchors.json \
    --output report_email.html

Output:
  - HTML file (self-contained, image embedded inline)
  - PDF (via WeasyPrint, 0.6in margins) unless --no-pdf
"""
import argparse
import base64
import json
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))
from build_adjustment_report import load_providers_tsv, load_providers_excel, load_market_anchors, build_exhibit

# ─── Brand palette (matches the matplotlib exhibit charts) ───────────────
NAVY = "#00485F"
BLUE = "#2F7DB8"
ORANGE = "#F0832E"
LIGHT_BLUE_BG = "#EAF2F8"
LIGHT_ORANGE_BG = "#FDF0E4"
GREY_TEXT = "#4A4A4A"
BORDER = "#E1E6EA"
ROW_ALT = "#F7F9FB"


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def money(v):
    return f"${v:,.0f}"


def build_html(providers, anchors, exhibit_path, meta):
    total_current = sum(p['current_base'] for p in providers)
    total_planned = sum(p['planned_base'] for p in providers)
    total_proposed = sum(p['proposed_base'] for p in providers)
    delta_current = total_proposed - total_current
    delta_planned = total_proposed - total_planned
    pct_current = (delta_current / total_current * 100) if total_current else 0
    pct_planned = (delta_planned / total_planned * 100) if total_planned else 0
    p50 = anchors.get('p50', 0)
    avg_proposed = total_proposed / len(providers) if providers else 0
    vs_median_pct = ((avg_proposed - p50) / p50 * 100) if p50 else 0

    img_tag = ""
    if exhibit_path and Path(exhibit_path).exists():
        b64 = base64.b64encode(Path(exhibit_path).read_bytes()).decode("ascii")
        img_tag = (
            f'<img src="data:image/png;base64,{b64}" alt="Market band exhibit" '
            f'style="width:100%;max-width:100%;height:auto;border-radius:6px;border:1px solid {BORDER};display:block;">'
        )

    rows_html = ""
    for i, p in enumerate(providers):
        dc = p['proposed_base'] - p['current_base']
        dp = p['proposed_base'] - p['planned_base']
        bg = ROW_ALT if i % 2 else "#FFFFFF"
        rows_html += f"""
        <tr style="background:{bg};">
          <td style="padding:10px 14px;border-bottom:1px solid {BORDER};font-weight:600;color:#1A1A1A;">{esc(p['name'])}</td>
          <td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:center;color:{GREY_TEXT};">{p.get('yoe','N/A')}</td>
          <td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:right;color:{GREY_TEXT};">{money(p['current_base'])}</td>
          <td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:right;color:{GREY_TEXT};">{money(p['planned_base'])}</td>
          <td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:right;font-weight:600;color:{NAVY};">{money(p['proposed_base'])}</td>
          <td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:right;color:{ORANGE};font-weight:600;">+{money(dc)}</td>
          <td style="padding:10px 14px;border-bottom:1px solid {BORDER};text-align:right;color:{ORANGE};">+{money(dp)}</td>
        </tr>"""

    rationale_items = meta.get('rationale') or [
        "Internal equity: proposed salaries maintain YOE-based progression without inversions.",
        "External anchoring: proposed bases position providers within the p25–p75 market band.",
        "Targeted adjustment: focuses on mid-career providers where compression risk is highest.",
        "Compression guardrails: no cascade adjustments proposed at this time.",
    ]
    governance_items = meta.get('governance_notes') or [
        "Proposed salaries do not create senior-junior inversions.",
        "All percentiles are from governance reference (post-annual-increase file).",
        "Percentiles are NOT FTE-adjusted — part-time lower percentile positioning is correct.",
        "Benchmark source: blended weighted-average of aggregated survey data (weighted by sample size).",
    ]

    def bullet_list(items, color):
        html = ""
        for it in items:
            html += f"""
            <tr>
              <td style="padding:4px 0 4px 0;vertical-align:top;width:22px;color:{color};font-weight:700;">&#8226;</td>
              <td style="padding:4px 0;color:{GREY_TEXT};font-size:14px;line-height:1.5;">{esc(it)}</td>
            </tr>"""
        return html

    rationale_html = bullet_list(rationale_items, BLUE)
    governance_html = bullet_list(governance_items, NAVY)

    def kpi_card(label, value, sub, bg, fg):
        return f"""
        <td style="padding:0 6px;" width="25%">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:{bg};border-radius:8px;">
            <tr><td style="padding:16px 14px;">
              <div style="font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:{fg};opacity:.75;font-weight:700;">{label}</div>
              <div style="font-size:22px;font-weight:700;color:{fg};margin-top:4px;">{value}</div>
              <div style="font-size:12px;color:{fg};opacity:.7;margin-top:2px;">{sub}</div>
            </td></tr>
          </table>
        </td>"""

    kpi_row = "".join([
        kpi_card("Total Current Base", money(total_current), f"{len(providers)} provider(s)", LIGHT_BLUE_BG, NAVY),
        kpi_card("Total Proposed Base", money(total_proposed), "post-adjustment", LIGHT_ORANGE_BG, ORANGE),
        kpi_card("Incremental vs Current", f"+{money(delta_current)}", f"{pct_current:+.1f}%", LIGHT_BLUE_BG, NAVY),
        kpi_card("Incremental vs Planned", f"+{money(delta_planned)}", f"{pct_planned:+.1f}%", LIGHT_ORANGE_BG, ORANGE),
    ])

    generated_date = meta.get('snapshot_date', datetime.now().strftime('%Y-%m-%d'))
    cohort_name = esc(meta.get('cohort_name', 'N/A'))

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Fair Market Value Review — {cohort_name}</title>
<style>
  @page {{ size: Letter; margin: 0.6in; }}
  body {{ margin:0; padding:0; background:#EDF1F4; font-family: Carlito, 'Liberation Sans', -apple-system, Segoe UI, Helvetica, Arial, sans-serif; }}
  a {{ color: {BLUE}; }}
</style>
</head>
<body>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#EDF1F4;padding:28px 0;">
<tr><td align="center">
<table role="presentation" width="700" cellpadding="0" cellspacing="0"
       style="width:700px;max-width:700px;background:#FFFFFF;border-radius:10px;overflow:hidden;box-shadow:0 2px 14px rgba(0,0,0,0.08);">

  <!-- Header -->
  <tr><td style="background:linear-gradient(135deg,{NAVY} 0%,{BLUE} 100%);padding:36px 40px;">
    <div style="color:#FFFFFF;font-size:13px;letter-spacing:.08em;text-transform:uppercase;opacity:.85;font-weight:600;">Fair Market Value Review</div>
    <div style="color:#FFFFFF;font-size:28px;font-weight:700;margin-top:6px;">{cohort_name}</div>
    <div style="color:#FFFFFF;font-size:13px;opacity:.8;margin-top:8px;">Prepared {generated_date} &middot; Confidential — Compensation Committee Use Only</div>
  </td></tr>

  <!-- KPI cards -->
  <tr><td style="padding:28px 34px 10px 34px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>{kpi_row}</tr></table>
  </td></tr>

  <!-- Market positioning line -->
  <tr><td style="padding:6px 40px 22px 40px;">
    <div style="font-size:13px;color:{GREY_TEXT};">
      Average proposed base (<strong>{money(avg_proposed)}</strong>) vs market p50
      (<strong>{money(p50)}</strong>): <strong style="color:{ORANGE};">{vs_median_pct:+.1f}%</strong> vs market median.
    </div>
  </td></tr>

  <!-- Cohort table -->
  <tr><td style="padding:0 34px 8px 34px;">
    <div style="font-size:16px;font-weight:700;color:{NAVY};margin-bottom:10px;">Cohort Detail</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid {BORDER};border-radius:6px;overflow:hidden;">
      <tr style="background:{NAVY};">
        <th style="padding:10px 14px;text-align:left;color:#FFFFFF;font-size:12px;text-transform:uppercase;letter-spacing:.04em;">Name</th>
        <th style="padding:10px 14px;text-align:center;color:#FFFFFF;font-size:12px;text-transform:uppercase;letter-spacing:.04em;">YOE</th>
        <th style="padding:10px 14px;text-align:right;color:#FFFFFF;font-size:12px;text-transform:uppercase;letter-spacing:.04em;">Current</th>
        <th style="padding:10px 14px;text-align:right;color:#FFFFFF;font-size:12px;text-transform:uppercase;letter-spacing:.04em;">Planned</th>
        <th style="padding:10px 14px;text-align:right;color:#FFFFFF;font-size:12px;text-transform:uppercase;letter-spacing:.04em;">Proposed</th>
        <th style="padding:10px 10px;text-align:right;color:#FFFFFF;font-size:11px;text-transform:uppercase;letter-spacing:.03em;white-space:nowrap;">&Delta; Current</th>
        <th style="padding:10px 10px;text-align:right;color:#FFFFFF;font-size:11px;text-transform:uppercase;letter-spacing:.03em;white-space:nowrap;">&Delta; Planned</th>
      </tr>
      {rows_html}
    </table>
  </td></tr>

  <!-- Exhibit -->
  <tr><td style="padding:26px 34px 8px 34px;">
    <div style="font-size:16px;font-weight:700;color:{NAVY};margin-bottom:10px;">Market Range Exhibit</div>
    {img_tag}
    <div style="font-size:11px;color:#9AA5AD;margin-top:8px;font-style:italic;">
      Market band: p25–p75 (shaded). p50 median (line). p90 (triangle). Current (circle), Planned (square), Proposed (diamond).
    </div>
  </td></tr>

  <!-- Rationale -->
  <tr><td style="padding:26px 34px 8px 34px;">
    <div style="font-size:16px;font-weight:700;color:{NAVY};margin-bottom:6px;">Logic &amp; Empirical Rationale</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0">{rationale_html}</table>
  </td></tr>

  <!-- Governance -->
  <tr><td style="padding:18px 34px 8px 34px;">
    <div style="font-size:16px;font-weight:700;color:{NAVY};margin-bottom:6px;">Governance Notes &amp; Guardrails</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0">{governance_html}</table>
  </td></tr>

  <!-- Footer -->
  <tr><td style="padding:28px 34px 32px 34px;border-top:1px solid {BORDER};margin-top:20px;">
    <div style="font-size:11px;color:#9AA5AD;line-height:1.6;">
      Market percentiles are from published survey data; interpolated estimates are labeled as directional.
      Cost figures reflect base salary only unless otherwise noted. All data sourced from the most recent
      annual salary increases file and survey combined file. Generated {generated_date}.
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>"""
    return html


def main():
    parser = argparse.ArgumentParser(description="Build an HTML (email-ready) Compensation Adjustment Report")
    parser.add_argument('--config', help='JSON config file')
    parser.add_argument('--name', help='Cohort name')
    parser.add_argument('--providers-file', help='TSV or Excel file with provider data')
    parser.add_argument('--market-anchors-file', help='JSON file with p25/p50/p75/p90')
    parser.add_argument('--output', default='report_email.html')
    parser.add_argument('--no-pdf', action='store_true', help='Skip PDF export (requires WeasyPrint)')

    args = parser.parse_args()
    cli_output = args.output if '--output' in sys.argv else None
    if args.config:
        with open(args.config) as f:
            cfg = json.load(f)
        for k, v in cfg.items():
            setattr(args, k, v)
    # CLI --output wins over the config's output; and since configs are often
    # shared with build_adjustment_report.py (whose output is a .docx), coerce
    # a non-HTML output path to .html rather than writing HTML into a .docx.
    if cli_output:
        args.output = cli_output
    if not str(args.output).lower().endswith('.html'):
        args.output = str(Path(args.output).with_suffix('')) + '_email.html'

    if not args.providers_file:
        print("ERROR: --providers-file or --config required", file=sys.stderr)
        sys.exit(1)

    if args.providers_file.endswith('.xlsx'):
        providers = load_providers_excel(args.providers_file)
    else:
        providers = load_providers_tsv(args.providers_file)
    print(f"Loaded {len(providers)} providers from {args.providers_file}")

    anchors = {}
    if args.market_anchors_file:
        anchors = load_market_anchors(args.market_anchors_file)

    exhibit_path = None
    if anchors and providers:
        exhibit_path = str(Path(args.output).with_suffix('')) + '_exhibit.png'
        build_exhibit(providers, anchors, exhibit_path, title=f"Base Salary vs Market Band — {args.name or 'Cohort'}")

    meta = {
        'cohort_name': args.name or 'N/A',
        'snapshot_date': datetime.now().strftime('%Y-%m-%d'),
        'rationale': getattr(args, 'rationale', None),
        'governance_notes': getattr(args, 'governance_notes', None),
    }

    html = build_html(providers, anchors, exhibit_path, meta)
    Path(args.output).write_text(html, encoding='utf-8')
    print(f"HTML saved: {args.output}")

    if not args.no_pdf:
        try:
            import weasyprint
            pdf_path = str(Path(args.output).with_suffix('.pdf'))
            weasyprint.HTML(string=html, base_url=str(Path(args.output).parent)).write_pdf(pdf_path)
            print(f"PDF saved: {pdf_path} (0.6in margins)")
        except ImportError:
            print("WARNING: weasyprint not installed -- skipping PDF export. "
                  "Install with: pip install weasyprint", file=sys.stderr)


if __name__ == '__main__':
    main()
