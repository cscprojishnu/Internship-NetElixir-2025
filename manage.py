import fitz  # PyMuPDF
import re

def clean_pdf(input_path, output_path):
    doc = fitz.open(input_path)
    new_doc = fitz.open()

    answer_pattern = re.compile(r"^Correct Answer\s*-\s*\w+", re.IGNORECASE)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")

        # Split by lines and filter out answer and explanation block
        lines = text.split("\n")
        clean_lines = []
        skip = False
        for line in lines:
            if answer_pattern.match(line.strip()):
                skip = True
                continue
            if skip:
                # Assume explanation ends with a table or blank line
                if line.strip() == "":
                    skip = False
                continue
            clean_lines.append(line)

        # Create a new page with the filtered text
        new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
        text_cleaned = "\n".join(clean_lines)
        new_page.insert_text((50, 50), text_cleaned, fontsize=11)

    new_doc.save(output_path)
    new_doc.close()
    doc.close()

# Example usage:
clean_pdf(r"C:\Users\cscpr\Downloads\NEET-PG-2012-Question-Paper-With-Solutions.pdf", "output_cleaned.pdf")
