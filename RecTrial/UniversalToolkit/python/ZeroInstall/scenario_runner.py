#!/usr/bin/env python3
from __future__ import annotations
VERSION="1.0.0"
import argparse,csv
from pathlib import Path
from safety_runtime import make_run_output,require_existing_file,write_run_logs

def parse_args():
 p=argparse.ArgumentParser(description="Run base/optimistic/conservative scenarios from input CSV Amount column.")
 p.add_argument("input_csv",nargs="?",type=Path,help="Input CSV path")
 p.add_argument("--sample",action="store_true",help="Run with synthetic sample data")
 return p.parse_args()

def pf(v):
 try:return float(str(v).replace(',','').replace('$',''))
 except:return None

def main():
 a=parse_args(); out=make_run_output("scenario_runner")
 try:
  src=out/"sample_input.csv" if a.sample else a.input_csv
  if a.sample: src.write_text("Amount\n100\n200\n",encoding="utf-8")
  else: require_existing_file(src,"input CSV")
  dst=out/"scenario_results.csv"; sc=[("base",0), ("optimistic",0.05), ("conservative",-0.05)]
  with src.open("r",encoding="utf-8-sig",newline="") as f: rows=list(csv.DictReader(f))
  with dst.open("w",encoding="utf-8",newline="") as f:
   w=csv.writer(f); w.writerow(["Scenario","Rows","BaseTotal","ScenarioTotal","Delta"])
   for n,p in sc:
    vals=[pf(r.get("Amount")) for r in rows]; vals=[v for v in vals if v is not None]; b=sum(vals); s=sum(v*(1+p) for v in vals)
    w.writerow([n,len(vals),f"{b:.2f}",f"{s:.2f}",f"{s-b:.2f}"])
  summary=f"Scenarios evaluated: {len(sc)}. Output: {dst.name}"; write_run_logs(out,summary,{"tool":"scenario_runner","scenarios":len(sc),"output":dst.name}); print(summary); print(f"Output folder: {out}")
 except SystemExit: raise
 except Exception:
  write_run_logs(out,"Run failed. Check run_log.json.",{"tool":"scenario_runner","status":"failed"}); raise SystemExit("Error: processing failed. See run_summary.txt in output folder.")

if __name__=='__main__': main()
