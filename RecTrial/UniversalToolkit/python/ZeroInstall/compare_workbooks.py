#!/usr/bin/env python3
from __future__ import annotations
VERSION="1.0.0"
import argparse,csv,zipfile,re
from pathlib import Path
import xml.etree.ElementTree as ET
from safety_runtime import make_run_output,require_existing_file,write_run_logs
NS_MAIN={"main":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}; NS_REL={"rel":"http://schemas.openxmlformats.org/package/2006/relationships"}

def parse_args():
 p=argparse.ArgumentParser(description="Compare two xlsx/xlsm workbooks and output cell-level diff CSV.")
 p.add_argument("left_workbook",nargs="?",type=Path,help="Left workbook path")
 p.add_argument("right_workbook",nargs="?",type=Path,help="Right workbook path")
 p.add_argument("--sample",action="store_true",help="Run in sample mode (help-only for this tool)")
 return p.parse_args()

def _shared(z):
 if "xl/sharedStrings.xml" not in z.namelist(): return []
 root=ET.fromstring(z.read("xl/sharedStrings.xml")); vals=[]
 for si in root.findall("main:si",NS_MAIN): vals.append("".join(n.text or "" for n in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")))
 return vals

def _smap(z):
 wb=ET.fromstring(z.read("xl/workbook.xml")); rels=ET.fromstring(z.read("xl/_rels/workbook.xml.rels")); rm={r.attrib["Id"]:r.attrib["Target"] for r in rels.findall("rel:Relationship",NS_REL)}
 m={}
 for s in wb.find("main:sheets",NS_MAIN): m[s.attrib.get("name","")]="xl/"+rm.get(s.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",""),"").lstrip("/")
 return m

def _cells(z,p,sh):
 if p not in z.namelist(): return {}
 root=ET.fromstring(z.read(p)); out={}
 for c in root.findall('.//main:sheetData/main:row/main:c',NS_MAIN):
  ref=c.attrib.get('r','')
  if not re.match(r'^[A-Z]+\d+$',ref): continue
  t=c.attrib.get('t'); v=c.find('main:v',NS_MAIN); i=c.find('main:is',NS_MAIN)
  if i is not None: out[ref]="".join(n.text or "" for n in i.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')); continue
  if v is None or v.text is None: out[ref]=""; continue
  raw=v.text; out[ref]=sh[int(raw)] if t=='s' and raw.isdigit() and int(raw)<len(sh) else raw
 return out

def main():
 a=parse_args(); out=make_run_output("compare_workbooks")
 try:
  if a.sample: raise SystemExit("Error: sample mode for compare_workbooks is help-only; provide two workbook paths.")
  require_existing_file(a.left_workbook,"left workbook"); require_existing_file(a.right_workbook,"right workbook")
  with zipfile.ZipFile(a.left_workbook) as zl, zipfile.ZipFile(a.right_workbook) as zr:
   lsh,rsh=_shared(zl),_shared(zr); lm,rm=_smap(zl),_smap(zr); dif=[]
   for name in sorted(set(lm)|set(rm)):
    lc,rc=_cells(zl,lm.get(name,""),lsh),_cells(zr,rm.get(name,""),rsh)
    for ref in sorted(set(lc)|set(rc)):
      if lc.get(ref,"")!=rc.get(ref,""): dif.append((name,ref,lc.get(ref,""),rc.get(ref,"")))
  dst=out/"workbook_diffs.csv"
  with dst.open("w",encoding="utf-8",newline="") as f: w=csv.writer(f); w.writerow(["sheet","cell","left_value","right_value"]); w.writerows(dif)
  summary=f"Diff rows: {len(dif)}. Output: {dst.name}"; write_run_logs(out,summary,{"tool":"compare_workbooks","diff_rows":len(dif),"output":dst.name}); print(summary); print(f"Output folder: {out}")
 except SystemExit: raise
 except Exception:
  write_run_logs(out,"Run failed. Check run_log.json.",{"tool":"compare_workbooks","status":"failed"}); raise SystemExit("Error: processing failed. See run_summary.txt in output folder.")

if __name__=='__main__': main()
