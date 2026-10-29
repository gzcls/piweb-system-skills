from __future__ import annotations
import argparse, os, shutil, subprocess, tempfile
from pathlib import Path

def find_known(kind):
    if kind=='libreoffice': return shutil.which('soffice') or next((p for p in [r'C:\Program Files\LibreOffice\program\soffice.exe',r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'] if Path(p).is_file()), None)
    if kind=='word': return shutil.which('winword') or next((p for p in [r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',r'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE'] if Path(p).is_file()), None)
    if kind=='wps': return shutil.which('wps') or next((p for p in [r'C:\Program Files\Kingsoft\WPS Office\ksolaunch.exe',r'C:\Program Files (x86)\Kingsoft\WPS Office\ksolaunch.exe'] if Path(p).is_file()), None)

def main():
    ap=argparse.ArgumentParser(description='Report available DOCX renderers and optionally run LibreOffice.')
    ap.add_argument('docx',type=Path); ap.add_argument('--out',type=Path,default=Path('render-matrix')); ap.add_argument('--libreoffice',action='store_true'); ap.add_argument('--word-exe'); ap.add_argument('--wps-exe')
    a=ap.parse_args(); a.out.mkdir(parents=True,exist_ok=True); rows=[]
    for name,explicit in [('LibreOffice',find_known('libreoffice')),('Microsoft Word',a.word_exe or find_known('word')),('WPS',a.wps_exe or find_known('wps'))]:
        row={'renderer':name,'executable':explicit,'status':'available' if explicit else 'unavailable'}
        if name=='LibreOffice' and explicit and a.libreoffice:
            dest=a.out/a.docx.with_suffix('.pdf').name; profile=Path(tempfile.mkdtemp(prefix='docx-render-'))
            try:
                uri=profile.resolve().as_uri(); cmd=[explicit,'--headless','--norestore','--nolockcheck','--nodefault',f'-env:UserInstallation={uri}','--convert-to','pdf','--outdir',str(a.out.resolve()),str(a.docx.resolve())]
                proc=subprocess.run(cmd,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=180)
                row.update(status='pass' if proc.returncode==0 and dest.exists() else 'fail',output=str(dest),returncode=proc.returncode,stderr=proc.stderr[-500:])
            finally: shutil.rmtree(profile,ignore_errors=True)
        elif name in ('Microsoft Word','WPS') and explicit:
            row['status']='detected_only'; row['note']='No safe headless PDF contract is assumed; pass --word-exe/--wps-exe to an approved automation wrapper.'
        rows.append(row)
    print({'source':str(a.docx.resolve()),'renderers':rows})
    return 0 if all(r['status'] not in ('fail',) for r in rows) else 1
if __name__=='__main__': raise SystemExit(main())
