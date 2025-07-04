import fitz 
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

def extract_rna_info_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    overall_data = []
    region_data = []

    for page in doc:
        text = page.get_text()

        
        sample_blocks = re.split(r"Overall Results for sample\s+\d+\s*:", text)
        sample_ids = re.findall(r"Overall Results for sample\s+\d+\s*:\s*(.+)", text)

        for idx, block in enumerate(sample_blocks[1:]):
            sample_id = sample_ids[idx].strip() if idx < len(sample_ids) else f"Sample_{idx+1}"

          
            rna_conc_match = re.search(r"RNA Concentration:\s*([\d,\.]+)\s*(p|n)g\s*/\s*[μuµ]l", block, flags=re.IGNORECASE)
            if rna_conc_match:
                conc_value = float(rna_conc_match.group(1).replace(",", ""))
                unit = rna_conc_match.group(2).lower()
                rna_conc = conc_value / 1000 if unit == "p" else conc_value
            else:
                rna_conc = None

           
            rin_match = re.search(r"RNA Integrity Number \(RIN\):\s*([\d\.]+)", block)
            rin = float(rin_match.group(1)) if rin_match else None

            overall_data.append({
                "Index": idx,
                "Sample ID": sample_id,
                "RNA Concentration (ng/µl)": rna_conc,
                "RIN": rin
            })

    
        matches = re.findall(
            r"Region table for sample\s+\d+\s+:\s+(.+?)\s+Name\s+From\s+\[nt\].*?DV200\s+(\d+)\s+[\d,]+\s+[\d,\.]+\s+([\d,\.]+)",
            text,
            flags=re.DOTALL
        )

        for match in matches:
            sample_id = match[0].strip()
            percent_total = float(match[2].replace(",", ""))
            region_data.append({
                "Sample ID": sample_id,
                "% of Total/DV200": percent_total
            })

    
    df_overall = pd.DataFrame(overall_data)
    df_region = pd.DataFrame(region_data)
    final_df = pd.merge(df_overall, df_region, on="Sample ID", how="left")
    final_df.sort_values("Index", inplace=True)
    final_df.drop(columns="Index", inplace=True)
    return final_df

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        entry_pdf_path.delete(0, tk.END)
        entry_pdf_path.insert(0, file_path)

def browse_output_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_output_folder.delete(0, tk.END)
        entry_output_folder.insert(0, folder_path)

def drop(event):
    if event.data.endswith(".pdf"):
        entry_pdf_path.delete(0, tk.END)
        entry_pdf_path.insert(0, event.data.strip("{}"))
    else:
        messagebox.showerror("Invalid File", "Please drop a PDF file.")

def run_extraction():
    pdf_path = entry_pdf_path.get()
    output_folder = entry_output_folder.get()

    if not os.path.isfile(pdf_path):
        messagebox.showerror("Error", "Invalid PDF file selected.")
        return

    if not os.path.isdir(output_folder):
        messagebox.showerror("Error", "Invalid output folder selected.")
        return

    df = extract_rna_info_from_pdf(pdf_path)
    if not df.empty:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = os.path.join(output_folder, f"{base_name}_RNA_Summary.xlsx")
        df.to_excel(output_path, index=False)
        messagebox.showinfo("Success", f"Data saved to:\n{output_path}")
    else:
        messagebox.showwarning("No Data", "No RNA QC data found in the PDF.")


root = TkinterDnD.Tk()
root.title("RNA Summary Extractor (BioA)")
root.geometry("500x250")
root.config(bg="#f0f4f8")

font_label = ("Segoe UI", 10)
font_button = ("Segoe UI", 10, "bold")


tk.Label(root, text="PDF File:", bg="#f0f4f8", font=font_label).pack(pady=(10, 0))
entry_pdf_path = tk.Entry(root, width=60)
entry_pdf_path.pack(pady=2)
entry_pdf_path.drop_target_register(DND_FILES)
entry_pdf_path.dnd_bind('<<Drop>>', drop)
tk.Button(root, text="Browse PDF", command=browse_file, font=font_button).pack(pady=5)


tk.Label(root, text="Output Folder:", bg="#f0f4f8", font=font_label).pack(pady=(10, 0))
entry_output_folder = tk.Entry(root, width=60)
entry_output_folder.pack(pady=2)
tk.Button(root, text="Select Folder", command=browse_output_folder, font=font_button).pack(pady=5)


tk.Button(root, text="Extract and Save", command=run_extraction,
          bg="#4CAF50", fg="white", font=font_button).pack(pady=15)

root.mainloop()
