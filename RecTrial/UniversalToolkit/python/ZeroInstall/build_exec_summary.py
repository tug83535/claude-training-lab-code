#!/usr/bin/env python3
from __future__ import annotations
VERSION="1.0.0"
import argparse,csv
from pathlib import Path
from statistics import mean
from safety_runtime import make_run_output,require_existing_file,write_run_logs

def parse_args():
 p=argparse.ArgumentParser(description="Create plain-English executive summary from CSV numeric data.")
 p.add_argument("input_csv",nargs="?",type=Path,help="Input CSV path")
 p.add_argument("--sample",action="store_true",help="Run with synthetic sample data")
 return p.parse_args()

def pf(v):
 try:return float(str(v).replace(',','').replace('$',''))
 except:return None

def main():
 a=parse_args(); out=make_run_output("build_exec_summary")
 try:
  src=out/"sample_input.csv" if a.sample else a.input_csv
  if a.sample: src.write_text("Department,Amount\nA,100\nB,250\n",encoding="utf-8")
  else: require_existing_file(src,"input CSV")
  with src.open("r",encoding="utf-8-sig",newline="") as f: rows=list(csv.DictReader(f))
  vals=[pf(r.get("Amount")) for r in rows]; vals=[v for v in vals if v is not None]
  if not vals: raise SystemExit("Error: no numeric data detected in Amount column")
  md=out/"executive_summary.md"
  txt="\n".join(["# Executive Summary",f"- Rows analyzed: **{len(rows)}**",f"- Total: **${sum(vals):,.0f}**",f"- Average: **${mean(vals):,.0f}**"])
  md.write_text(txt+"\n",encoding="utf-8")
  summary=f"Executive summary generated: {md.name}"; write_run_logs(out,summary,{"tool":"build_exec_summary","rows":len(rows),"output":md.name}); print(summary); print(f"Output folder: {out}")
 except SystemExit: raise
 except Exception:
  write_run_logs(out,"Run failed. Check run_log.json.",{"tool":"build_exec_summary","status":"failed"}); raise SystemExit("Error: processing failed. See run_summary.txt in output folder.")

if __name__=='__main__': main()
