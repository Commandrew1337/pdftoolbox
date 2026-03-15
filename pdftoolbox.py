# -*- coding: utf-8 -*-
"""
PDF Toolbox GUI — One window for common PDF tasks

Includes tabs for:
  • Merge PDFs from a folder (natural sort) — Save As picker (no list)
  • Extract selected pages to a new PDF
  • Remove selected pages and save as new PDF
  • Insert one PDF into another (beginning/end/before/after page N)
  • Extract images from a PDF (requires PyMuPDF/fitz; tab disables if missing)
  • Convert PDF text to paragraphs and save as .txt
  • Unlock a password-protected PDF (copy pages to a new, unencrypted file)

Dependencies:
  - PyPDF2 (required)
  - PyMuPDF / fitz (optional; only for the Extract Images tab)
"""
from __future__ import annotations

# --- stdlib ---
import os, re, platform, subprocess
from datetime import datetime
from pathlib import Path

# --- tkinter ---
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- PDF libs ---
try:
    from PyPDF2 import PdfReader, PdfWriter, PdfMerger
except Exception:
    raise SystemExit("PyPDF2 is required. Install with: pip install PyPDF2")

try:
    import fitz  # PyMuPDF
    _FITZ_AVAILABLE = True
except Exception:
    fitz = None
    _FITZ_AVAILABLE = False

APP_TITLE = "PDF Toolbox"
VERSION = "1.0"

# ---------------- Utilities ----------------

def ts_for_filename() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def natural_key(s: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]

# ---------------- Page range parsing ----------------

def parse_page_selection_extract(selection: str, total_pages: int):
    if not selection:
        raise ValueError("No pages entered.")
    pages = set()
    tokens = [t.strip() for t in selection.replace(";", ",").split(",") if t.strip()]
    range_pattern = re.compile(r"^(\d+)\s*-\s*(\d+)$")
    num_pattern = re.compile(r"^\d+$")
    for tok in tokens:
        m = range_pattern.match(tok)
        if m:
            start, end = int(m.group(1)), int(m.group(2))
            if start > end:
                raise ValueError(f"Invalid range '{tok}': start > end.")
            for p in range(start, end + 1):
                if p < 1 or p > total_pages:
                    raise ValueError(f"Page {p} is out of bounds (1–{total_pages}).")
                pages.add(p - 1)
        elif num_pattern.match(tok):
            p = int(tok)
            if p < 1 or p > total_pages:
                raise ValueError(f"Page {p} is out of bounds (1–{total_pages}).")
            pages.add(p - 1)
        else:
            raise ValueError(f"Unrecognized token '{tok}'. Use formats like 1,3,5-7.")
    return sorted(pages)

def parse_page_selection_remove(selection: str, total_pages: int):
    pages_to_remove = set()
    selection = (selection or "").strip()
    if not selection:
        return []
    parts = [p.strip() for p in selection.replace(";", ",").split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            start_str, end_str = [x.strip() for x in part.split("-", 1)]
            start = int(start_str); end = int(end_str)
            if start < 1 or end < 1: raise ValueError("Pages must be >= 1")
            if start > end: raise ValueError("Range start greater than end")
            for p in range(start, end + 1):
                if p > total_pages: raise ValueError(f"Page {p} exceeds document length ({total_pages})")
                pages_to_remove.add(p - 1)
        else:
            p = int(part)
            if p < 1: raise ValueError("Pages must be >= 1")
            if p > total_pages: raise ValueError(f"Page {p} exceeds document length ({total_pages})")
            pages_to_remove.add(p - 1)
    return sorted(pages_to_remove)

# ---------------- Text extraction helpers ----------------

def extract_text_from_pdf(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    return "\n\n".join((page.extract_text() or "") for page in reader.pages)

def reflow_paragraphs(raw_text: str) -> str:
    text = raw_text.replace("\r\n", "\n").replace("\r", "\n").expandtabs(4)
    lines = text.split("\n"); paragraphs = []; buf = []
    import re as _re
    def flush():
        if not buf: return
        parts=[]
        for i,line in enumerate(buf):
            line=line.strip();
            if not line: continue
            if i==0: parts.append(line)
            else:
                prev = parts[-1]
                parts[-1] = (prev[:-1] + line) if prev.endswith('-') else (prev + ' ' + line)
        para=_re.sub(r"\s+"," "," ".join(parts)).strip()
        if para: paragraphs.append(para)
        buf.clear()
    for line in lines:
        if line.strip()=="": flush()
        else: buf.append(line)
    flush(); return "\n\n".join(paragraphs)

# ---------------- Base tab ----------------
class BaseTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.status_var = tk.StringVar(value="Ready.")
        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate")
    def set_status(self, text: str): self.status_var.set(text)
    def set_progress_mode(self, mode: str):
        self.progress.config(mode=mode)
        (self.progress.start(10) if mode=="indeterminate" else self.progress.stop())
    def set_progress(self, value: int, maximum: int|None=None):
        if maximum is not None: self.progress.config(maximum=maximum)
        self.progress['value']=value

# ---------------- Merge (no listbox; Save As) ----------------
class MergeTab(BaseTab):
    def __init__(self, master):
        super().__init__(master)
        self.folder = tk.StringVar()
        self.output = tk.StringVar()
        self._build()
    def _build(self):
        pad={"padx":10,"pady":8}
        frm=ttk.Frame(self); frm.grid(row=0,column=0,sticky="nsew",**pad)
        self.columnconfigure(0, weight=1)
        ttk.Label(frm,text="Folder with PDFs:").grid(row=0,column=0,sticky="w")
        ttk.Entry(frm,textvariable=self.folder).grid(row=0,column=1,sticky="ew",padx=(6,6))
        frm.columnconfigure(1,weight=1)
        ttk.Button(frm,text="Browse…",command=self._browse_folder).grid(row=0,column=2)
        ttk.Label(frm,text="Save merged PDF As:").grid(row=1,column=0,sticky="w")
        ttk.Entry(frm,textvariable=self.output).grid(row=1,column=1,sticky="ew",padx=(6,6))
        ttk.Button(frm,text="Choose…",command=self._browse_output).grid(row=1,column=2)
        status=ttk.Frame(self); status.grid(row=1,column=0,sticky="ew",padx=10,pady=(0,8))
        ttk.Label(status,textvariable=self.status_var).pack(anchor="w")
        self.progress.pack(in_=status,fill="x",pady=(4,0))
        ttk.Button(self,text="Merge PDFs",command=self.merge).grid(row=2,column=0,sticky="e",padx=10,pady=(0,10))
    def _browse_folder(self):
        path=filedialog.askdirectory(title="Select folder containing PDFs to merge")
        if path: self.folder.set(path)
    def _browse_output(self):
        folder=Path(self.folder.get().strip()) if self.folder.get().strip() else Path.home()
        default=folder/f"combined_{ts_for_filename()}.pdf"
        path=filedialog.asksaveasfilename(title="Save merged PDF as…",initialdir=str(folder),initialfile=default.name,defaultextension=".pdf",filetypes=[("PDF files","*.pdf")])
        if path: self.output.set(path)
    def merge(self):
        folder=Path(self.folder.get().strip()); out=self.output.get().strip()
        if not folder.exists() or not folder.is_dir():
            messagebox.showerror("Invalid folder","Please select a valid folder containing PDFs."); return
        if not out:
            messagebox.showerror("Missing output","Please choose where to save the merged PDF."); return
        pdfs=[p for p in folder.iterdir() if p.is_file() and p.suffix.lower()==".pdf"]
        pdfs=[p for p in pdfs if not p.name.startswith("combined_")]
        pdfs=sorted(pdfs,key=lambda p: natural_key(p.name))
        if not pdfs:
            messagebox.showerror("No PDFs","No PDF files found in the selected folder."); return
        merger=PdfMerger()
        try:
            self.set_status("Merging…"); self.set_progress_mode("determinate"); self.set_progress(0,len(pdfs))
            for i,p in enumerate(pdfs, start=1):
                try: merger.append(str(p))
                except Exception: pass
                self.set_progress(i)
            self.set_status("Writing output…"); merger.write(out); merger.close()
            self.set_status(f"Done → {out}"); messagebox.showinfo("Merge complete",f"Saved to:\n{out}")
        except Exception as e:
            try: merger.close()
            except Exception: pass
            messagebox.showerror("Merge failed", str(e))

# ---------------- Extract Pages ----------------
class ExtractPagesTab(BaseTab):
    def __init__(self, master):
        super().__init__(master)
        self.src=tk.StringVar(); self.pages=tk.StringVar(); self.dst=tk.StringVar(); self.page_count=tk.IntVar(value=0)
        self._build()
    def _build(self):
        pad={"padx":10,"pady":8}
        grid=ttk.Frame(self); grid.grid(row=0,column=0,sticky="nsew",**pad); self.columnconfigure(0,weight=1)
        ttk.Label(grid,text="Source PDF:").grid(row=0,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.src).grid(row=0,column=1,sticky="ew",padx=(6,6)); grid.columnconfigure(1,weight=1)
        ttk.Button(grid,text="Browse…",command=self._browse_src).grid(row=0,column=2)
        ttk.Label(grid,text="Pages (1-based): e.g. 1,3,5-7").grid(row=1,column=0,sticky="w",columnspan=3)
        ttk.Entry(grid,textvariable=self.pages).grid(row=2,column=0,columnspan=2,sticky="ew")
        ttk.Label(grid,textvariable=self.page_count).grid(row=2,column=2,sticky="e")
        ttk.Label(grid,text="Save As:").grid(row=3,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.dst).grid(row=3,column=1,sticky="ew",padx=(6,6))
        ttk.Button(grid,text="Choose…",command=self._browse_dst).grid(row=3,column=2)
        status=ttk.Frame(self); status.grid(row=1,column=0,sticky="ew",padx=10)
        ttk.Label(status,textvariable=self.status_var).pack(anchor="w"); self.progress.pack(in_=status,fill="x",pady=(4,0))
        ttk.Button(self,text="Extract",command=self.extract).grid(row=2,column=0,sticky="e",padx=10,pady=(0,10))
    def _browse_src(self):
        path=filedialog.askopenfilename(title="Select PDF",filetypes=[("PDF files","*.pdf")])
        if path:
            self.src.set(path)
            try:
                reader=PdfReader(path); self.page_count.set(len(reader.pages))
                name=Path(path).stem+f"_extracted_{ts_for_filename()}.pdf"; self.dst.set(str(Path(path).with_name(name)))
            except Exception as e: messagebox.showerror("Error",f"Failed to read PDF: {e}")
    def _browse_dst(self):
        path=filedialog.asksaveasfilename(title="Save extracted pages as…",defaultextension=".pdf",filetypes=[("PDF files","*.pdf")])
        if path: self.dst.set(path)
    def extract(self):
        src,dst,sel=self.src.get().strip(),self.dst.get().strip(),self.pages.get().strip()
        if not src or not os.path.isfile(src): messagebox.showerror("Missing file","Select a valid source PDF."); return
        if not dst: messagebox.showerror("Missing output","Choose where to save the output PDF."); return
        try: reader=PdfReader(src); total=len(reader.pages)
        except Exception as e: messagebox.showerror("Error",f"Failed to read PDF: {e}"); return
        try: indices=parse_page_selection_extract(sel,total)
        except ValueError as ve: messagebox.showerror("Invalid input",str(ve)); return
        writer=PdfWriter(); self.set_progress_mode("determinate"); self.set_progress(0,len(indices))
        for i,idx in enumerate(indices, start=1): writer.add_page(reader.pages[idx]); self.set_progress(i)
        try:
            with open(dst,"wb") as f: writer.write(f)
            self.set_status(f"Saved {len(indices)} page(s) → {dst}"); messagebox.showinfo("Done",f"Saved to:\n{dst}")
        except Exception as e: messagebox.showerror("Error",f"Failed to write: {e}")

# ---------------- Remove Pages ----------------
class RemovePagesTab(BaseTab):
    def __init__(self, master):
        super().__init__(master)
        self.src=tk.StringVar(); self.pages=tk.StringVar(); self.dst=tk.StringVar(); self.page_count=tk.IntVar(value=0)
        self._build()
    def _build(self):
        pad={"padx":10,"pady":8}
        grid=ttk.Frame(self); grid.grid(row=0,column=0,sticky="nsew",**pad); self.columnconfigure(0,weight=1)
        ttk.Label(grid,text="Source PDF:").grid(row=0,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.src).grid(row=0,column=1,sticky="ew",padx=(6,6)); grid.columnconfigure(1,weight=1)
        ttk.Button(grid,text="Browse…",command=self._browse_src).grid(row=0,column=2)
        ttk.Label(grid,text="Pages to REMOVE (1-based): 2,4-6,10").grid(row=1,column=0,sticky="w",columnspan=3)
        ttk.Entry(grid,textvariable=self.pages).grid(row=2,column=0,columnspan=2,sticky="ew")
        ttk.Label(grid,textvariable=self.page_count).grid(row=2,column=2,sticky="e")
        ttk.Label(grid,text="Save As:").grid(row=3,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.dst).grid(row=3,column=1,sticky="ew",padx=(6,6))
        ttk.Button(grid,text="Choose…",command=self._browse_dst).grid(row=3,column=2)
        status=ttk.Frame(self); status.grid(row=1,column=0,sticky="ew",padx=10)
        ttk.Label(status,textvariable=self.status_var).pack(anchor="w"); self.progress.pack(in_=status,fill="x",pady=(4,0))
        ttk.Button(self,text="Remove",command=self.remove).grid(row=2,column=0,sticky="e",padx=10,pady=(0,10))
    def _browse_src(self):
        path=filedialog.askopenfilename(title="Select PDF",filetypes=[("PDF files","*.pdf")])
        if path:
            self.src.set(path)
            try:
                reader=PdfReader(path); self.page_count.set(len(reader.pages))
                name=Path(path).stem+f"_removed_{ts_for_filename()}.pdf"; self.dst.set(str(Path(path).with_name(name)))
            except Exception as e: messagebox.showerror("Error",f"Failed to read PDF: {e}")
    def _browse_dst(self):
        path=filedialog.asksaveasfilename(title="Save output as…",defaultextension=".pdf",filetypes=[("PDF files","*.pdf")])
        if path: self.dst.set(path)
    def remove(self):
        src,dst,sel=self.src.get().strip(),self.dst.get().strip(),self.pages.get().strip()
        if not src or not os.path.isfile(src): messagebox.showerror("Missing file","Select a valid source PDF."); return
        if not dst: messagebox.showerror("Missing output","Choose where to save the output PDF."); return
        try: reader=PdfReader(src); total=len(reader.pages)
        except Exception as e: messagebox.showerror("Error",f"Failed to read PDF: {e}"); return
        try: to_remove=set(parse_page_selection_remove(sel,total))
        except ValueError as ve: messagebox.showerror("Invalid input",str(ve)); return
        writer=PdfWriter(); self.set_progress_mode("determinate"); self.set_progress(0,total)
        kept=0
        for i in range(total):
            if i not in to_remove: writer.add_page(reader.pages[i]); kept+=1
            self.set_progress(i+1)
        try:
            with open(dst,"wb") as f: writer.write(f)
            self.set_status(f"Removed {len(to_remove)} page(s). Remaining {kept}. → {dst}"); messagebox.showinfo("Done",f"Saved to:\n{dst}")
        except Exception as e: messagebox.showerror("Error",f"Failed to write: {e}")

# ---------------- Insert PDF ----------------
class InsertTab(BaseTab):
    def __init__(self, master):
        super().__init__(master)
        self.base=tk.StringVar(); self.content=tk.StringVar(); self.dst=tk.StringVar();
        self.mode=tk.StringVar(value="At beginning"); self.page_n=tk.IntVar(value=1); self.base_pages=tk.IntVar(value=0)
        self._build()
    def _build(self):
        pad={"padx":10,"pady":8}
        grid=ttk.Frame(self); grid.grid(row=0,column=0,sticky="nsew",**pad); self.columnconfigure(0,weight=1)
        ttk.Label(grid,text="Content PDF (to insert):").grid(row=0,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.content).grid(row=0,column=1,sticky="ew",padx=(6,6)); grid.columnconfigure(1,weight=1)
        ttk.Button(grid,text="Browse…",command=self._browse_content).grid(row=0,column=2)
        ttk.Label(grid,text="Base PDF (insert into):").grid(row=1,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.base).grid(row=1,column=1,sticky="ew",padx=(6,6))
        ttk.Button(grid,text="Browse…",command=self._browse_base).grid(row=1,column=2)
        info=ttk.Frame(grid); info.grid(row=2,column=0,columnspan=3,sticky="ew")
        ttk.Label(info,text="Base pages:").pack(side="left"); ttk.Label(info,textvariable=self.base_pages).pack(side="left",padx=(6,0))
        pos=ttk.Frame(grid); pos.grid(row=3,column=0,columnspan=3,sticky="ew")
        ttk.Label(pos,text="Insert position:").pack(side="left")
        cmb=ttk.Combobox(pos,textvariable=self.mode,state="readonly",values=["At beginning","Before page…","After page…","At end"],width=18)
        cmb.pack(side="left",padx=(6,12)); cmb.bind("<<ComboboxSelected>>", lambda e: self._update_page_input())
        ttk.Label(pos,text="Page number (1-based):").pack(side="left")
        self.spn=tk.Spinbox(pos,from_=1,to=1,textvariable=self.page_n,width=6,state="disabled"); self.spn.pack(side="left",padx=(6,0))
        ttk.Label(grid,text="Save As:").grid(row=4,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.dst).grid(row=4,column=1,sticky="ew",padx=(6,6))
        ttk.Button(grid,text="Choose…",command=self._browse_dst).grid(row=4,column=2)
        status=ttk.Frame(self); status.grid(row=1,column=0,sticky="ew",padx=10)
        ttk.Label(status,textvariable=self.status_var).pack(anchor="w"); self.progress.pack(in_=status,fill="x",pady=(4,0))
        ttk.Button(self,text="Merge / Insert",command=self.merge).grid(row=2,column=0,sticky="e",padx=10,pady=(0,10))
    def _update_page_input(self): self.spn.config(state=("normal" if self.mode.get() in ("Before page…","After page…") else "disabled"))
    def _browse_content(self):
        path=filedialog.askopenfilename(title="Select Content PDF",filetypes=[("PDF files","*.pdf")])
        if path: self.content.set(path); self._suggest_output()
    def _browse_base(self):
        path=filedialog.askopenfilename(title="Select Base PDF",filetypes=[("PDF files","*.pdf")])
        if path:
            self.base.set(path)
            try: reader=PdfReader(path); count=len(reader.pages)
            except Exception: count=0
            self.base_pages.set(count)
            try: self.spn.config(to=max(1,count))
            except Exception: pass
            if count>=1: self.page_n.set(1)
            self._suggest_output()
    def _browse_dst(self):
        path=filedialog.asksaveasfilename(title="Save merged PDF as…",defaultextension=".pdf",filetypes=[("PDF files","*.pdf")])
        if path: self.dst.set(path)
    def _suggest_output(self):
        base=self.base.get().strip(); content=self.content.get().strip()
        if not base or not content: return
        base_name=Path(base).stem; content_name=Path(content).stem; mode=self.mode.get()
        tag = "beginning" if mode=="At beginning" else ("end" if mode=="At end" else (f"before_{self.page_n.get()}" if mode=="Before page…" else f"after_{self.page_n.get()}"))
        directory=Path(base).parent; out=directory/f"{base_name}__with__{content_name}_{tag}.pdf"
        if not self.dst.get().strip(): self.dst.set(str(out))
    @staticmethod
    def _insert_index(base_count:int, mode:str, page_1based:int)->int:
        if mode=="At beginning": return 0
        if mode=="At end": return base_count
        if base_count<=0: raise ValueError("Base PDF has no pages.")
        if page_1based<1 or page_1based>base_count: raise ValueError(f"Page number must be between 1 and {base_count} for the base PDF.")
        if mode=="Before page…": return page_1based-1
        if mode=="After page…": return page_1based
        raise ValueError("Unknown position mode.")
    def merge(self):
        base,content,dst=self.base.get().strip(),self.content.get().strip(),self.dst.get().strip()
        if not content or not os.path.isfile(content): messagebox.showerror("Missing file","Select a valid Content PDF."); return
        if not base or not os.path.isfile(base): messagebox.showerror("Missing file","Select a valid Base PDF."); return
        if not dst: messagebox.showerror("Missing output","Choose where to save the output PDF."); return
        try:
            base_r=PdfReader(base); content_r=PdfReader(content); base_count=len(base_r.pages); idx=self._insert_index(base_count,self.mode.get(),self.page_n.get())
        except Exception as e: messagebox.showerror("Error",str(e)); return
        writer=PdfWriter(); self.set_progress_mode("determinate"); self.set_progress(0, base_count+len(content_r.pages)); cur=0
        for i in range(idx): writer.add_page(base_r.pages[i]); cur+=1; self.set_progress(cur)
        for p in content_r.pages: writer.add_page(p); cur+=1; self.set_progress(cur)
        for i in range(idx, len(base_r.pages)): writer.add_page(base_r.pages[i]); cur+=1; self.set_progress(cur)
        try:
            with open(dst,"wb") as f: writer.write(f)
            self.set_status(f"Done → {dst}"); messagebox.showinfo("Success",f"Merged PDF saved:\n{dst}")
        except Exception as e: messagebox.showerror("Error",f"Failed to write: {e}")

# ---------------- Extract Images ----------------
class ImagesTab(BaseTab):
    def __init__(self, master):
        super().__init__(master); self.src=tk.StringVar(); self.outdir=tk.StringVar(); self._build()
    def _build(self):
        pad={"padx":10,"pady":8}
        grid=ttk.Frame(self); grid.grid(row=0,column=0,sticky="nsew",**pad); self.columnconfigure(0,weight=1)
        if not _FITZ_AVAILABLE:
            ttk.Label(grid,text=("PyMuPDF (fitz) is not installed. Install with:\n    pip install pymupdf\n\nThis tab will remain disabled until fitz is available."),foreground="#a00").grid(row=0,column=0,sticky="w"); return
        ttk.Label(grid,text="Source PDF:").grid(row=0,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.src).grid(row=0,column=1,sticky="ew",padx=(6,6)); grid.columnconfigure(1,weight=1)
        ttk.Button(grid,text="Browse…",command=self._browse_src).grid(row=0,column=2)
        ttk.Label(grid,text="Output folder:").grid(row=1,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.outdir).grid(row=1,column=1,sticky="ew",padx=(6,6))
        ttk.Button(grid,text="Choose…",command=self._browse_outdir).grid(row=1,column=2)
        status=ttk.Frame(self); status.grid(row=1,column=0,sticky="ew",padx=10)
        ttk.Label(status,textvariable=self.status_var).pack(anchor="w"); self.progress.pack(in_=status,fill="x",pady=(4,0))
        ttk.Button(self,text="Extract images",command=self.extract).grid(row=2,column=0,sticky="e",padx=10,pady=(0,10))
    def _browse_src(self):
        path=filedialog.askopenfilename(title="Select PDF",filetypes=[("PDF files","*.pdf")])
        if path:
            self.src.set(path); pdf_name=Path(path).stem; out=Path(path).with_name(f"{pdf_name}_images"); self.outdir.set(str(out))
    def _browse_outdir(self):
        d=filedialog.askdirectory(title="Choose output folder")
        if d: self.outdir.set(d)
    def extract(self):
        if not _FITZ_AVAILABLE: messagebox.showerror("Missing dependency","PyMuPDF (fitz) is required for image extraction."); return
        src,outdir=self.src.get().strip(),self.outdir.get().strip()
        if not src or not os.path.isfile(src): messagebox.showerror("Missing file","Select a valid source PDF."); return
        if not outdir: messagebox.showerror("Missing output folder","Choose an output folder."); return
        try:
            os.makedirs(outdir,exist_ok=True); doc=fitz.open(src); total=len(doc)
            self.set_progress_mode("determinate"); self.set_progress(0,total); count=0
            for page_index in range(total):
                page=doc.load_page(page_index); images=page.get_images(full=True)
                for img_index,img in enumerate(images):
                    xref=img[0]; base_image=doc.extract_image(xref); image_bytes=base_image["image"]; image_ext=base_image["ext"]
                    image_path=Path(outdir)/f"page{page_index+1}_img{img_index+1}.{image_ext}"
                    with open(image_path,"wb") as f: f.write(image_bytes); count+=1
                self.set_progress(page_index+1)
            self.set_status(f"Extracted {count} image(s) → {outdir}"); messagebox.showinfo("Done",f"Extracted {count} image(s) to:\n{outdir}")
        except Exception as e: messagebox.showerror("Error",f"Failed: {e}")

# ---------------- PDF → Text ----------------
class TextTab(BaseTab):
    def __init__(self, master):
        super().__init__(master); self.src=tk.StringVar(); self.dst=tk.StringVar(); self._build()
    def _build(self):
        pad={"padx":10,"pady":8}
        grid=ttk.Frame(self); grid.grid(row=0,column=0,sticky="nsew",**pad); self.columnconfigure(0,weight=1)
        ttk.Label(grid,text="Source PDF:").grid(row=0,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.src).grid(row=0,column=1,sticky="ew",padx=(6,6)); grid.columnconfigure(1,weight=1)
        ttk.Button(grid,text="Browse…",command=self._browse_src).grid(row=0,column=2)
        ttk.Label(grid,text="Save text As:").grid(row=1,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.dst).grid(row=1,column=1,sticky="ew",padx=(6,6))
        ttk.Button(grid,text="Choose…",command=self._browse_dst).grid(row=1,column=2)
        status=ttk.Frame(self); status.grid(row=1,column=0,sticky="ew",padx=10)
        ttk.Label(status,textvariable=self.status_var).pack(anchor="w"); self.progress.pack(in_=status,fill="x",pady=(4,0))
        ttk.Button(self,text="Extract text",command=self.run).grid(row=2,column=0,sticky="e",padx=10,pady=(0,10))
    def _browse_src(self):
        path=filedialog.askopenfilename(title="Select PDF",filetypes=[("PDF files","*.pdf")])
        if path:
            self.src.set(path); suggested=Path(path).with_suffix("").name+f"_text_{ts_for_filename()}.txt"; self.dst.set(str(Path(path).with_name(suggested)))
    def _browse_dst(self):
        path=filedialog.asksaveasfilename(title="Save text as…",defaultextension=".txt",filetypes=[("Text files","*.txt"),("All files","*.*")])
        if path: self.dst.set(path)
    def run(self):
        src,dst=self.src.get().strip(),self.dst.get().strip()
        if not src or not os.path.isfile(src): messagebox.showerror("Missing file","Select a valid source PDF."); return
        if not dst: messagebox.showerror("Missing output","Choose where to save the .txt file."); return
        try:
            self.set_status("Extracting raw text…"); raw=extract_text_from_pdf(src)
            if not raw.strip(): messagebox.showwarning("No text","No extractable text found (PDF may be scanned images). "); return
            self.set_status("Reflowing paragraphs…"); text=reflow_paragraphs(raw)
            with open(dst,"w",encoding="utf-8") as f: f.write(text)
            self.set_status(f"Done → {dst}"); messagebox.showinfo("Done",f"Text saved to:\n{dst}")
        except Exception as e: messagebox.showerror("Error",f"Failed: {e}")

# ---------------- Unlock ----------------
class UnlockTab(BaseTab):
    def __init__(self, master):
        super().__init__(master); self.src=tk.StringVar(); self.dst=tk.StringVar(); self.password=tk.StringVar(); self.show_pass=tk.BooleanVar(value=False); self._build()
    def _build(self):
        pad={"padx":10,"pady":8}
        grid=ttk.Frame(self); grid.grid(row=0,column=0,sticky="nsew",**pad); self.columnconfigure(0,weight=1)
        ttk.Label(grid,text="Locked PDF:").grid(row=0,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.src).grid(row=0,column=1,sticky="ew",padx=(6,6)); grid.columnconfigure(1,weight=1)
        ttk.Button(grid,text="Browse…",command=self._browse_src).grid(row=0,column=2)
        row2=ttk.Frame(grid); row2.grid(row=1,column=0,columnspan=3,sticky="ew")
        ttk.Label(row2,text="Password:").pack(side="left")
        self.pass_entry=ttk.Entry(row2,textvariable=self.password,show="*"); self.pass_entry.pack(side="left",fill="x",expand=True,padx=(6,6))
        ttk.Checkbutton(row2,text="Show",variable=self.show_pass,command=self._toggle_show).pack(side="left")
        ttk.Label(grid,text="Save unlocked As:").grid(row=2,column=0,sticky="w")
        ttk.Entry(grid,textvariable=self.dst).grid(row=2,column=1,sticky="ew",padx=(6,6))
        ttk.Button(grid,text="Choose…",command=self._browse_dst).grid(row=2,column=2)
        status=ttk.Frame(self); status.grid(row=1,column=0,sticky="ew",padx=10)
        ttk.Label(status,textvariable=self.status_var).pack(anchor="w"); self.progress.pack(in_=status,fill="x",pady=(4,0))
        ttk.Button(self,text="Unlock",command=self.unlock).grid(row=2,column=0,sticky="e",padx=10,pady=(0,10))
    def _toggle_show(self): self.pass_entry.config(show=("" if self.show_pass.get() else "*"))
    def _browse_src(self):
        path=filedialog.askopenfilename(title="Select password-protected PDF",filetypes=[("PDF files","*.pdf")])
        if path:
            self.src.set(path); base_dir=Path(path).parent; base_name=Path(path).stem; self.dst.set(str(base_dir/f"{base_name}_unlocked.pdf"))
    def _browse_dst(self):
        path=filedialog.asksaveasfilename(title="Save unlocked PDF as…",defaultextension=".pdf",filetypes=[("PDF files","*.pdf")])
        if path: self.dst.set(path)
    def unlock(self):
        src,dst,pwd=self.src.get().strip(),self.dst.get().strip(),self.password.get()
        if not src or not os.path.isfile(src): messagebox.showerror("Missing file","Select a valid locked PDF."); return
        if not dst: messagebox.showerror("Missing output","Choose where to save the unlocked PDF."); return
        try:
            self.set_progress_mode("indeterminate")
            with open(src,"rb") as f:
                reader=PdfReader(f)
                if reader.is_encrypted:
                    result=reader.decrypt(pwd or "")
                    if not result: raise ValueError("Incorrect password or decryption failed.")
                total=len(reader.pages); writer=PdfWriter(); self.set_progress_mode("determinate"); self.set_progress(0,total)
                for i,page in enumerate(reader.pages,start=1): writer.add_page(page); self.set_progress(i)
                with open(dst,"wb") as out: writer.write(out)
            self.set_status(f"Unlocked → {dst}"); messagebox.showinfo("Success",f"Unlocked PDF saved to:\n{dst}")
        except Exception as e: messagebox.showerror("Error",str(e))

# ---------------- Main app ----------------
class PDFToolboxApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} v{VERSION}")
        # Build UI first so Tk can compute natural/requested sizes
        self._build()
        # --- Auto-size to the smallest needed dimensions ---
        self.update_idletasks()  # let geometry settle
        req_w = self.winfo_reqwidth()
        req_h = self.winfo_reqheight()
        # Set geometry to requested size and lock it as the minimum
        self.geometry(f'{req_w}x{req_h}')
        self.minsize(req_w, req_h)
        # Optional: keep resizable; user can expand if they want
        self.resizable(True, True)

    def _build(self):
        nb=ttk.Notebook(self); nb.pack(fill="both",expand=True)
        nb.add(MergeTab(nb), text="Merge")
        nb.add(ExtractPagesTab(nb), text="Extract Pages")
        nb.add(RemovePagesTab(nb), text="Remove Pages")
        nb.add(InsertTab(nb), text="Insert PDF")
        nb.add(ImagesTab(nb), text="Extract Images")
        nb.add(TextTab(nb), text="PDF → Text")
        nb.add(UnlockTab(nb), text="Unlock")

def main():
    PDFToolboxApp().mainloop()

if __name__=="__main__":
    main()
