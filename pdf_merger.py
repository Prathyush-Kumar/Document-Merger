import os
import time
import tempfile
from tkinter import Tk, filedialog, simpledialog, messagebox
from pypdf import PdfWriter
from docx2pdf import convert

# Save location
SAVE_PATH = r"C:\Users\Admin\Documents"


def select_files(root):
    files = filedialog.askopenfilenames(
        title="Select DOCX and PDF files",
        filetypes=[("Documents", "*.docx *.pdf")]
    )
    return list(files)


def get_filename(root):
    name = simpledialog.askstring(
        "File Name",
        "Enter merged file name (without .pdf):",
        parent=root
    )
    return name


def merge_files(files, filename):
    writer = PdfWriter()
    temp_files = []

    try:
        for file in files:
            try:
                if file.lower().endswith(".docx"):
                    print(f"Converting: {file}")

                    # Use a temp file to avoid overwriting existing PDFs
                    temp_pdf = tempfile.mktemp(suffix=".pdf")
                    convert(file, temp_pdf)
                    time.sleep(1)  # Prevents Word crash

                    temp_files.append(temp_pdf)

                    with open(temp_pdf, "rb") as f:
                        writer.append(f)

                elif file.lower().endswith(".pdf"):
                    print(f"Adding PDF: {file}")

                    with open(file, "rb") as f:
                        writer.append(f)

            except Exception as e:
                print(f"Error processing {file}: {e}")

        # Ensure save directory exists
        os.makedirs(SAVE_PATH, exist_ok=True)

        output = os.path.join(SAVE_PATH, f"{filename}.pdf")

        with open(output, "wb") as f:
            writer.write(f)

        print(f"\n✅ Merged file created: {output}")
        return output

    finally:
        # Always clean up temp files
        for temp in temp_files:
            if os.path.exists(temp):
                os.remove(temp)
                print(f"🧹 Cleaned up temp file: {temp}")


# Run Program
root = Tk()
root.withdraw()

try:
    files = select_files(root)

    if not files:
        print("No files selected.")
    else:
        filename = get_filename(root)

        if not filename:
            print("No filename entered.")
        else:
            output_path = merge_files(files, filename)

            messagebox.showinfo(
                "Done",
                f"Merged PDF saved to:\n{output_path}"
            )
finally:
    root.destroy()
