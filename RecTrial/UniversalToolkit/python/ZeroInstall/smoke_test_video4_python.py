#!/usr/bin/env python3
from __future__ import annotations
import subprocess,sys,tempfile
from pathlib import Path
VERSION="1.0.0"

def ok(cmd): return subprocess.run(cmd,capture_output=True,text=True).returncode==0

def main():
 base=Path(__file__).resolve().parent; py=sys.executable; res=[]
 with tempfile.TemporaryDirectory() as td:
  d=Path(td); inp=d/'input.csv'; inp.write_text('Department,Amount,Actual,Baseline\nA,100,110,100\nB,200,180,200\n',encoding='utf-8')
  res.append(ok([py,str(base/'sanitize_dataset.py'),str(inp)]))
  res.append(ok([py,str(base/'variance_classifier.py'),str(inp)]))
  res.append(ok([py,str(base/'scenario_runner.py'),str(inp)]))
  res.append(ok([py,str(base/'build_exec_summary.py'),str(inp)]))
  res.append(ok([py,str(base/'compare_workbooks.py'),'--help']) and ok([py,str(base/'sheets_to_csv.py'),'--help']))
 p=sum(1 for x in res if x); t=len(res); print(f"Smoke results: {p}/{t} PASS");
 if p!=t: raise SystemExit(1)

if __name__=='__main__': main()
