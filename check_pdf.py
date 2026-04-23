"""Quick PDF diagnostic — run this first to check what's in the PDF."""
import sys
import pdfplumber

pdf_path = sys.argv[1] if len(sys.argv) > 1 else input("PDF path: ").strip('"')

with pdfplumber.open(pdf_path) as pdf:
    print(f"Pages: {len(pdf.pages)}")
    for i, page in enumerate(pdf.pages):
        words = page.extract_words()
        text  = page.extract_text() or ""
        imgs  = page.images
        print(f"\nPage {i+1}: {len(words)} words, {len(imgs)} images, {len(text)} chars")
        if words:
            print("First 10 words:", [w['text'] for w in words[:10]])
        elif imgs:
            print("  → Page contains images only (scanned PDF — needs OCR)")
        else:
            print("  → Page is completely empty")
