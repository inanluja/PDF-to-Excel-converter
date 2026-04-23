"""Print every row from the PDF so we can see the exact security name text."""
import sys
import pdfplumber

pdf_path = sys.argv[1] if len(sys.argv) > 1 else input("PDF path: ").strip('"')

def group_rows(words, tol=5):
    if not words:
        return []
    rows, cur, cy = [], [words[0]], words[0]["top"]
    for w in words[1:]:
        if abs(w["top"] - cy) <= tol:
            cur.append(w)
        else:
            rows.append(sorted(cur, key=lambda x: x["x0"]))
            cur, cy = [w], w["top"]
    rows.append(sorted(cur, key=lambda x: x["x0"]))
    return rows

keywords = ["SPDR", "VANGUARD", "ISHARES", "iSHARES", "GOLD", "TREASURY",
            "MID-CAP", "S&P", "ETF", "TRUST", "BOND", "CORP"]

with pdfplumber.open(pdf_path) as pdf:
    for pi, page in enumerate(pdf.pages):
        words = page.extract_words(x_tolerance=3, y_tolerance=3)
        words = sorted(words, key=lambda w: (round(w["top"]/3)*3, w["x0"]))
        rows = group_rows(words)
        print(f"\n=== Page {pi+1}: {len(rows)} rows ===")
        for i, row in enumerate(rows):
            text = " ".join(w["text"] for w in row)
            # Print rows that contain financial keywords OR numbers > 1000
            has_keyword = any(k.upper() in text.upper() for k in keywords)
            has_number  = any(
                c.isdigit() for c in text
            )
            if has_keyword:
                print(f"  [{i:03d}] >>> {text[:150]}")
            elif i < 5:  # also print first 5 rows for context
                print(f"  [{i:03d}]     {text[:150]}")
