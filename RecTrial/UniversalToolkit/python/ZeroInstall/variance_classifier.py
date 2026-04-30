#!/usr/bin/env python3
from __future__ import annotations
VERSION="1.0.0"
import argparse,csv
from pathlib import Path
from safety_runtime import make_run_output,require_existing_file,write_run_logs

def parse_args():
 p=argparse.ArgumentParser(description="Classify Actual vs Baseline variance into Direction/Materiality.")
 p.add_argument("input_csv",nargs="?",type=Path,help="Input CSV path")
 p.add_argument("--sample",action="store_true",help="Run with synthetic sample data")
 return p.parse_args()

def pf(v):
 try:return float(str(v).replace(',','').replace('$',''))
 except:return None

def main():
 a=parse_args(); out=make_run_output("variance_classifier")
 try:
  src=out/"sample_input.csv" if a.sample else a.input_csv
  if a.sample: src.write_text("Actual,Baseline\n110,100\n180,200\n",encoding="utf-8")
  else: require_existing_file(src,"input CSV")
  dst=out/"variance_classified.csv"
  with src.open("r",encoding="utf-8-sig",newline="") as f: rows=list(csv.DictReader(f))
  if not rows: raise SystemExit("Error: input CSV has no rows")
  fns=list(rows[0].keys())+["Variance","VariancePct","Direction","Materiality"]
  with dst.open("w",encoding="utf-8",newline="") as f:
   w=csv.DictWriter(f,fieldnames=fns); w.writeheader(); cnt=0
   for r in rows:
    act,base=pf(r.get("Actual")),pf(r.get("Baseline"))
    if act is None or base is None: d,m,delta,pct="unknown","insufficient_data",None,None
    else:
      delta=act-base; pct=None if base==0 else delta/base
      d="favorable" if delta>=0 else "unfavorable"; m="material" if abs(delta)>=1000 or (pct is not None and abs(pct)>=0.1) else "non_material"
    o=dict(r); o.update({"Variance":"" if delta is None else f"{delta:.2f}","VariancePct":"" if pct is None else f"{pct:.4f}","Direction":d,"Materiality":m}); w.writerow(o); cnt+=1
  summary=f"Classified rows: {cnt}. Output: {dst.name}"; write_run_logs(out,summary,{"tool":"variance_classifier","rows":cnt,"output":dst.name}); print(summary); print(f"Output folder: {out}")
 except SystemExit: raise
 except Exception:
  write_run_logs(out,"Run failed. Check run_log.json.",{"tool":"variance_classifier","status":"failed"}); raise SystemExit("Error: processing failed. See run_summary.txt in output folder.")

if __name__=='__main__': main()
