import os
import zipfile
import PyPDF2
import pandas as pd
import tempfile
import shutil
from pathlib import Path

def parse_zip_pdfs():
    # input/output paths

    input_folder = Path.home() / "Documents/pdf_input"
    output_folder = Path.home() / "Documents/pdf_output"
    os.makedirs(output_folder, exist_ok=True)

    
    zip_files = list(input_folder.glob("*.zip"))
    if not zip_files:
        print(f" No ZIP file found in {input_folder}")
        return

    print(f"Found {len(zip_files)} ZIP file(s) to process")


    for zip_path in zip_files:
        print(f"\n Processing: {zip_path.name}")
        
        zip_name = zip_path.stem  
        output_excel = output_folder / f"{zip_name}_pdf_summary.xlsx"

        
        temp_dir = tempfile.mkdtemp()

        try:
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(temp_dir)

            data = []

            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if file.lower().endswith(".pdf"):
                        file_path = os.path.join(root, file)
                        rel_path = os.path.relpath(file_path, temp_dir)

                        try:
                            with open(file_path, "rb") as pdf_file:
                                reader = PyPDF2.PdfReader(pdf_file)
                                num_pages = len(reader.pages)

                            data.append([
                                os.path.splitext(file)[0],  # Tag
                                "",                         # Designation
                                num_pages,                  # Page count
                                rel_path                    # Path in Folder
                            ])

                        except Exception as e:
                            data.append([os.path.splitext(file)[0], "", "Error", rel_path])

           
            if data:
                df = pd.DataFrame(data, columns=["Tag", "Designation", "Page", "Path in Folder"])
                df.to_excel(output_excel, index=False)
                print(f"Excel saved: {output_excel} ({len(data)} PDFs processed)")
            else:
                print(f"No PDFs found in {zip_path.name}")

        except Exception as e:
            print(f"Error processing {zip_path.name}: {str(e)}")
        finally:
            shutil.rmtree(temp_dir)

    print(f"\n All ZIP files processed! Check output folder: {output_folder}")


if __name__ == "__main__":
    parse_zip_pdfs()
