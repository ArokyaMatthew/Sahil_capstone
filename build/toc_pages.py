"""Reads the rendered report PDF and records the body page on which each chapter starts (for the TOC)."""
import json
import re
import subprocess
import sys

pdf, out = sys.argv[1], sys.argv[2]
n = int(re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout).group(1))
heads = ['ABSTRACT', 'INTRODUCTION', 'LITERATURE REVIEW', 'PROBLEM IDENTIFICATION & OBJECTIVES', 'EXISTING SYSTEM',
         'PROPOSED SYSTEM', 'SYSTEM ARCHITECTURE', 'TOOLS/TECHNOLOGIES USED', 'CONCLUSION', 'REFERENCES']
found, first_body = {}, None
for pg in range(1, n + 1):
    txt = subprocess.run(['pdftotext', '-f', str(pg), '-l', str(pg), '-layout', pdf, '-'], capture_output=True, text=True).stdout
    for k, h in enumerate(heads, 1):
        if str(k) not in found and re.search(rf'^\s*{k}\.\s+{re.escape(h)}\s*$', txt, re.M):
            found[str(k)] = pg
            if k == 1:
                first_body = pg
pages = {k: v - first_body + 1 for k, v in found.items()}
json.dump(pages, open(out, 'w'))
print(pages)
