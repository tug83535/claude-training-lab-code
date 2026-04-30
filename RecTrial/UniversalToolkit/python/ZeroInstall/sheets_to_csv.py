#!/usr/bin/env python3
from __future__ import annotations
VERSION="1.0.0"
import argparse,csv,zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
from safety_runtime import make_run_output,require_existing_file,write_run_logs
NS_MAIN={"main":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}; NS_REL={"rel":"http://schemas.openxmlformats.org/package/2006/relationships"}; NS_DOC_REL="{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

def parse_args():
 p=argparse.ArgumentParser(description="Extract selected sheets from workbook into CSV files in toolkit outputs/.")
 p.add_argument("workbook",nargs="?",type=Path,help="Workbook path")
 p.add_argument("--sheets",default="Sheet1",help="Comma-separated sheet names")
 p.add_argument("--sample",action="store_true",help="Run in sample mode (help-only for this tool)")
 return p.parse_args()

def _shared(z):
 if 'xl/sharedStrings.xml' not in z.namelist(): return []
 root=ET.fromstring(z.read('xl/sharedStrings.xml')); vals=[]
 for si in root.findall('main:si',NS_MAIN): vals.append(''.join(n.text or '' for n in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')))
 return vals

def _smap(z):
 wb=ET.fromstring(z.read('xl/workbook.xml')); rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels')); rm={i.attrib['Id']:i.attrib['Target'] for i in rels.findall('rel:Relationship',NS_REL)}
 res={}; sheets=wb.find('main:sheets',NS_MAIN)
 if sheets is None:return res
 for s in sheets:
  t=rm.get(s.attrib.get(NS_DOC_REL,''));
  if t: res[s.attrib.get('name','')] = 'xl/' + t.lstrip('/')
 return res

def _rows(z,p,sh):
 if p not in z.namelist(): return []
 root=ET.fromstring(z.read(p)); out=[]
 for r in root.findall('.//main:sheetData/main:row',NS_MAIN):
  row=[]
  for c in r.findall('main:c',NS_MAIN):
   t=c.attrib.get('t'); v=c.find('main:v',NS_MAIN); i=c.find('main:is',NS_MAIN)
   if i is not None: row.append(''.join(n.text or '' for n in i.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'))); continue
   if v is None or v.text is None: row.append(''); continue
   raw=v.text; row.append(sh[int(raw)] if t=='s' and raw.isdigit() and int(raw)<len(sh) else raw)
  out.append(row)
 return out

def main():
 a=parse_args(); out=make_run_output("sheets_to_csv")
 try:
  if a.sample: raise SystemExit("Error: sample mode for sheets_to_csv is help-only; provide workbook path.")
  require_existing_file(a.workbook,"workbook")
  with zipfile.ZipFile(a.workbook) as z:
   sh=_shared(z); mp=_smap(z); names=[x.strip() for x in a.sheets.split(',') if x.strip()]; created=[]
   for n in names:
    p=mp.get(n)
    if not p: continue
    rows=_rows(z,p,sh); dst=out/("".join(ch if ch.isalnum() or ch in '-_' else '_' for ch in n)+'.csv')
    with dst.open('w',encoding='utf-8',newline='') as f: csv.writer(f).writerows(rows)
    created.append(dst.name)
  summary=f"Extracted files: {len(created)}"; write_run_logs(out,summary,{"tool":"sheets_to_csv","files_created":created}); print(summary); print(f"Output folder: {out}")
 except SystemExit: raise
 except Exception:
  write_run_logs(out,"Run failed. Check run_log.json.",{"tool":"sheets_to_csv","status":"failed"}); raise SystemExit("Error: processing failed. See run_summary.txt in output folder.")

if __name__=='__main__': main()
