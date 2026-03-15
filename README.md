# pdftoolbox
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
