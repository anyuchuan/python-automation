"""独立渲染页面，供研究者核对文字层与实际外观。原PDF只读。"""
import argparse
from pathlib import Path
import fitz
from build import locate_pdf

def render(output):
 output=Path(output);output.mkdir(parents=True,exist_ok=True)
 with fitz.open(locate_pdf()) as d:
  for i,p in enumerate(d):
   (output/f'PDF第{i+1:02}页.jpg').write_bytes(p.get_pixmap(matrix=fitz.Matrix(1.6,1.6)).tobytes('jpg',jpg_quality=85))
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--output',type=Path,required=True);render(p.parse_args().output)
