#!/usr/bin/env python3
"""
FMV Request Register -- track every incoming review request in one place.

Compensation teams get these continuously (a candidate needs a range, a chief
wants an adjustment reviewed, a committee needs a packet). They usually live in
an inbox until someone asks "what's outstanding?" This keeps a simple auditable
register instead. Storage is a plain CSV so it opens in Excel, diffs in git, and
needs no database.

  python3 scripts/fmv_register.py add --provider "Provider 1" --type adjustment --due 2026-09-01
  python3 scripts/fmv_register.py update FMV-0001 --status in_review --tcc-percentile 75.3 \
      --wrvu-percentile 50.9 --alignment comp_above_production
  python3 scripts/fmv_register.py list --open
  python3 scripts/fmv_register.py summary
"""
import argparse, csv, sys
from datetime import datetime, date
from pathlib import Path

REGISTER_DIR = Path("register")
REGISTER_FILE = REGISTER_DIR / "fmv_requests.csv"
FIELDS = ["request_id","date_received","provider","specialty","review_type","requester",
          "due_date","status","current_base","proposed_base","tcc_percentile",
          "wrvu_percentile","alignment_flag","combined_n","decision","decision_date",
          "deliverables","notes","last_updated"]
REVIEW_TYPES = ["new_hire", "adjustment", "benchmark_only"]
STATUSES = ["intake","awaiting_data","in_review","pending_committee","approved","declined","withdrawn"]
OPEN_STATUSES = ["intake","awaiting_data","in_review","pending_committee"]


def _load():
    if not REGISTER_FILE.exists():
        return []
    with REGISTER_FILE.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _save(rows):
    REGISTER_DIR.mkdir(parents=True, exist_ok=True)
    with REGISTER_FILE.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS); w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})


def _next_id(rows):
    nums = []
    for r in rows:
        rid = r.get("request_id", "")
        if rid.startswith("FMV-"):
            try: nums.append(int(rid.split("-")[1]))
            except (IndexError, ValueError): pass
    return f"FMV-{max(nums) + 1 if nums else 1:04d}"


def cmd_add(args):
    rows = _load(); rid = _next_id(rows)
    row = {k: "" for k in FIELDS}
    row.update({"request_id": rid, "date_received": args.received or date.today().isoformat(),
                "provider": args.provider, "specialty": args.specialty or "",
                "review_type": args.type, "requester": args.requester or "",
                "due_date": args.due or "", "status": "intake",
                "current_base": args.current_base or "", "notes": args.notes or "",
                "last_updated": datetime.now().isoformat(timespec="seconds")})
    rows.append(row); _save(rows)
    print(f"Logged {rid}: {args.provider} ({args.type})" + (f", due {args.due}" if args.due else ""))
    print(f"Register: {REGISTER_FILE}")
    return rid


def cmd_update(args):
    rows = _load()
    match = [r for r in rows if r["request_id"].upper() == args.request_id.upper()]
    if not match:
        print(f"ERROR: no request {args.request_id!r} in {REGISTER_FILE}", file=sys.stderr)
        sys.exit(1)
    row = match[0]; changes = []
    for field, value in [("status",args.status),("proposed_base",args.proposed_base),
                         ("current_base",args.current_base),("tcc_percentile",args.tcc_percentile),
                         ("wrvu_percentile",args.wrvu_percentile),("alignment_flag",args.alignment),
                         ("combined_n",args.combined_n),("decision",args.decision),
                         ("deliverables",args.deliverables),("notes",args.notes),("due_date",args.due)]:
        if value is not None:
            row[field] = str(value); changes.append(field)
    if args.status in ("approved","declined","withdrawn") and not args.decision_date:
        row["decision_date"] = date.today().isoformat(); changes.append("decision_date")
    if args.decision_date:
        row["decision_date"] = args.decision_date; changes.append("decision_date")
    row["last_updated"] = datetime.now().isoformat(timespec="seconds")
    _save(rows)
    print(f"Updated {row['request_id']}: {', '.join(changes) or 'no changes'}")


def _fmt_row(r):
    pct = r.get("tcc_percentile","")
    pct_s = f"p{float(pct):.0f}" if pct else "-"
    flag = {"comp_above_production":"! comp>prod","production_above_comp":"! prod>comp",
            "aligned":"aligned"}.get(r.get("alignment_flag",""), "-")
    return (f"  {r['request_id']:<9} {r['status']:<17} {r['provider'][:22]:<22} "
            f"{r.get('review_type',''):<14} {pct_s:<5} {flag:<12} due:{r.get('due_date','') or '-'}")


def cmd_list(args):
    rows = _load()
    if not rows:
        print("No requests logged yet. Add one:\n"
              "  python3 scripts/fmv_register.py add --provider NAME --type adjustment")
        return
    if args.open: rows = [r for r in rows if r["status"] in OPEN_STATUSES]
    if args.status: rows = [r for r in rows if r["status"] == args.status]
    if args.type: rows = [r for r in rows if r["review_type"] == args.type]
    if not rows:
        print("No matching requests."); return
    rows.sort(key=lambda r: (r.get("due_date") or "9999", r["request_id"]))
    print(f"\n  {'ID':<9} {'STATUS':<17} {'PROVIDER':<22} {'TYPE':<14} {'TCC':<5} {'ALIGNMENT':<12} DUE")
    print("  " + "-" * 96)
    for r in rows: print(_fmt_row(r))
    print(f"\n  {len(rows)} request(s)\n")


def cmd_summary(args):
    rows = _load()
    if not rows:
        print("Register is empty."); return
    by_status, by_type, flagged, overdue = {}, {}, [], []
    today = date.today().isoformat()
    for r in rows:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
        by_type[r["review_type"]] = by_type.get(r["review_type"], 0) + 1
        if r.get("alignment_flag","").startswith(("comp_above","production_above")):
            flagged.append(r)
        if r.get("due_date") and r["due_date"] < today and r["status"] in OPEN_STATUSES:
            overdue.append(r)
    print(f"\nFMV REQUEST REGISTER — {len(rows)} total\n")
    print("  By status:")
    for k in STATUSES:
        if by_status.get(k): print(f"    {k:<18} {by_status[k]}")
    print("\n  By type:")
    for k, v in sorted(by_type.items()): print(f"    {k:<18} {v}")
    open_count = sum(by_status.get(s, 0) for s in OPEN_STATUSES)
    print(f"\n  Open: {open_count}   Closed: {len(rows) - open_count}")
    if overdue:
        print(f"\n  OVERDUE ({len(overdue)}):")
        for r in overdue: print(f"    {r['request_id']} {r['provider']} — due {r['due_date']}")
    if flagged:
        print(f"\n  Alignment flags to document ({len(flagged)}):")
        for r in flagged: print(f"    {r['request_id']} {r['provider']} — {r['alignment_flag']}")
    print()


def main():
    ap = argparse.ArgumentParser(description="FMV request register")
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("add"); a.add_argument("--provider", required=True)
    a.add_argument("--specialty"); a.add_argument("--type", choices=REVIEW_TYPES, required=True)
    a.add_argument("--requester"); a.add_argument("--due"); a.add_argument("--received")
    a.add_argument("--current-base"); a.add_argument("--notes"); a.set_defaults(func=cmd_add)
    u = sub.add_parser("update"); u.add_argument("request_id")
    u.add_argument("--status", choices=STATUSES); u.add_argument("--current-base")
    u.add_argument("--proposed-base"); u.add_argument("--tcc-percentile")
    u.add_argument("--wrvu-percentile"); u.add_argument("--alignment"); u.add_argument("--combined-n")
    u.add_argument("--decision"); u.add_argument("--decision-date"); u.add_argument("--deliverables")
    u.add_argument("--due"); u.add_argument("--notes"); u.set_defaults(func=cmd_update)
    l = sub.add_parser("list"); l.add_argument("--open", action="store_true")
    l.add_argument("--status", choices=STATUSES); l.add_argument("--type", choices=REVIEW_TYPES)
    l.set_defaults(func=cmd_list)
    s = sub.add_parser("summary"); s.set_defaults(func=cmd_summary)
    args = ap.parse_args(); args.func(args)


if __name__ == "__main__":
    main()
