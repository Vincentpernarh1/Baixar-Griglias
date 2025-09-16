import os
import sys
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from openpyxl import load_workbook, Workbook
from collections import defaultdict
import pandas as pd
import re
import threading

# Store all unique Volume TT values per model
volume_tt_set_by_model = defaultdict(set)

def extract_model_from_filename(filename):
    """Extract the first 3 characters from filename as model code"""
    base_name = os.path.basename(filename)
    return base_name[:3] if len(base_name) >= 3 else base_name

def collect_volume_tt(folder_path):
    search_pattern = os.path.join(folder_path, '*.xlsx')
    xlsx_files = glob.glob(search_pattern)
    for file in xlsx_files:
        try:
            wb = load_workbook(filename=file, read_only=False, data_only=True)
            if "Griglia Mondo - Volumi" in wb.sheetnames:
                ws = wb["Griglia Mondo - Volumi"]
            else:
                ws = wb["Grid World - Volume"]
            
            for row_dim in ws.row_dimensions.values():
                row_dim.hidden = False
            for col_dim in ws.column_dimensions.values():
                col_dim.hidden = False
            
            col_idx = 4
            model_code = str(ws.cell(row=12, column=4).value or "")
            
            while True:
                volume_head = ws.cell(row=15, column=col_idx).value
                if volume_head is None:
                    break
                col_idx += 1
            
            volume_tt = ws.cell(row=15, column=col_idx - 1).value
            if volume_tt:
                volume_tt_set_by_model[model_code].add(int(volume_tt))
        except Exception as e:
            print(f"Error reading TT in {file}: {e}")

def transform_and_expand(folder_path, De_Para_file_path, output_file_path):
    volume_tt_set_by_model.clear()  # Clear from previous runs
    collect_volume_tt(folder_path)
    volume_tt_sum_by_model = {
        model: sum(tts) for model, tts in volume_tt_set_by_model.items()
    }
    
    search_pattern = os.path.join(folder_path, '*.xlsx')
    xlsx_files = glob.glob(search_pattern)
    os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
    
    merged_rows = []
    header = [
        "Packet", "Code", "Concat", "Version", "Volume Head", "Volume", "Volume TT",
        "SINCOM", "Model code", "Model", "Concat", "Multivalues", "File_Model"
    ]
    
    for file in xlsx_files:
        try:
            # Extract model from filename (first 3 characters)
            file_model = extract_model_from_filename(file)
            print(f"Processing file: {os.path.basename(file)}, extracted model: {file_model}")
            
            wb = load_workbook(filename=file, read_only=False, data_only=True)
            if "Griglia Mondo - Volumi" in wb.sheetnames:
                ws = wb["Griglia Mondo - Volumi"]
            else:
                ws = wb["Grid World - Volume"]
            
            for row_dim in ws.row_dimensions.values():
                row_dim.hidden = False
            for col_dim in ws.column_dimensions.values():
                col_dim.hidden = False
            
            model_metadata = []
            col_idx = 4
            model_code = str(ws.cell(row=12, column=4).value or "")
            model = str(ws.cell(row=11, column=4).value or "")
            volume_tt_sum = volume_tt_sum_by_model.get(model_code, 0)
            
            while True:
                volume_head = ws.cell(row=15, column=col_idx).value
                if volume_head is None:
                    break
                
                version = ws.cell(row=13, column=col_idx).value
                sincom = ws.cell(row=14, column=col_idx).value
                
                metadata = {
                    "col_idx": col_idx,
                    "version": str(version or ""),
                    "volume_head": str(volume_head or ""),
                    "sincom": str(sincom or ""),
                    "model_code": model_code,
                    "model": model
                }
                model_metadata.append(metadata)
                col_idx += 1
            
            for row in range(11, 10001):
                if row == 12 or row == 13 or row == 14 or row == 15:
                    continue
                
                packet = ws.cell(row=row, column=1).value
                code = ws.cell(row=row, column=2).value
                
                if packet is None and code is None:
                    continue
                
                packet_str = str(packet or "")
                code_str = str(code or "").strip()
                concat = str(packet_str + code_str).strip()
                
                for meta in model_metadata:
                    col = meta["col_idx"]
                    volume_val = ws.cell(row=row, column=col).value
                    volume_str = str(volume_val or "")
                    engine_family = str(volume_val).strip() if volume_val else ""
                    
                    merged_rows.append([
                        packet_str,
                        code_str,
                        concat,
                        meta["version"],
                        meta["volume_head"],
                        volume_str,
                        str(volume_tt_sum),
                        meta["sincom"],
                        meta["model_code"],
                        meta["model"],
                        concat,
                        engine_family,
                        file_model  # Add the extracted file model
                    ])
            
            print(f"Processed: {file}")
        except Exception as e:
            print(f"Error processing {file}: {e}")
    
    # Create DataFrame from merged_rows for use in Update_De_Para
    df = pd.DataFrame(merged_rows, columns=header)
    
    # Save the expanded file
    final_wb = Workbook()
    final_ws = final_wb.active
    final_ws.title = "Expanded_Mapped"
    final_ws.append(header)
    for row in merged_rows:
        final_ws.append(row)
    final_wb.save(output_file_path)
    
    # Update De_Para file with the DataFrame
    if De_Para_file_path and os.path.exists(De_Para_file_path):
        Update_De_Para(De_Para_file_path, df)
    else:
        print("De_Para file not selected or doesn't exist. Skipping De_Para update.")
    
    print(f"Final file saved to: {output_file_path}")

def Update_De_Para(De_Para_file_path, df):
    try:
        # Load the existing De_Para file into a DataFrame
        De_Para = pd.read_excel(De_Para_file_path, sheet_name="Coded")
        
        # Remove duplicates from Griglia Italiano and Griglia Inglês columns
        De_Para = De_Para.drop_duplicates(subset=['Griglia Italiano', 'Griglia Inglês'])
        
        print(f"Processing {len(De_Para)} rows from De_Para file...")
        
        # Create a dictionary to store results by unique combination
        results_dict = {}
        
        # Get unique file models from the data
        unique_file_models = df['File_Model'].unique()
        print(f"Found file models: {list(unique_file_models)}")
        
        # Process each row in the De_Para file
        for index, row in De_Para.iterrows():
            italiano_value = str(row['Griglia Italiano']).strip() if pd.notna(row['Griglia Italiano']) else ""
            ingles_value = str(row['Griglia Inglês']).strip() if pd.notna(row['Griglia Inglês']) else ""
            
            print(f"Processing De_Para row {index + 1}: Italiano='{italiano_value}', Inglês='{ingles_value}'")
            
            # Process each file model
            for file_model in unique_file_models:
                model_df = df[df['File_Model'] == file_model]
                
                # Find matches for Italiano
                italiano_matches = pd.DataFrame()
                if italiano_value:
                    italiano_matches = model_df[model_df['Packet'].str.strip().str.lower() == italiano_value.lower()]
                
                # Find matches for Inglês
                ingles_matches = pd.DataFrame()
                if ingles_value:
                    ingles_matches = model_df[model_df['Packet'].str.strip().str.lower() == ingles_value.lower()]
                
                # Get unique multivalues from both matches
                all_multivalues = set()
                
                if not italiano_matches.empty:
                    all_multivalues.update(italiano_matches['Multivalues'].unique())
                
                if not ingles_matches.empty:
                    all_multivalues.update(ingles_matches['Multivalues'].unique())
                
                # Create entries for each unique multivalue
                for multivalue in all_multivalues:
                    if not multivalue or str(multivalue).strip() == '':
                        continue
                    
                    key = (multivalue, italiano_value, ingles_value, file_model)
                    
                    if key not in results_dict:
                        # Initialize the result row
                        results_dict[key] = {
                            'Multivalues': multivalue,
                            'Griglia Italiano': italiano_value,
                            'Griglia Inglês': ingles_value,
                            'Model': file_model,
                            'Resp.1': '',
                            'Resp.2': ''
                        }
                    
                    # Check if this multivalue appears in Italiano matches
                    italiano_has_multivalue = not italiano_matches.empty and \
                                            any(italiano_matches['Multivalues'] == multivalue)
                    
                    # Check if this multivalue appears in Inglês matches
                    ingles_has_multivalue = not ingles_matches.empty and \
                                          any(ingles_matches['Multivalues'] == multivalue)
                    
                    # Set Resp.1 if found in Italiano matches
                    if italiano_has_multivalue:
                        results_dict[key]['Resp.1'] = multivalue
                    
                    # Set Resp.2 if found in Inglês matches
                    if ingles_has_multivalue:
                        results_dict[key]['Resp.2'] = multivalue
        
        # Convert results dictionary to list
        flattened_results = list(results_dict.values())
        
        # Create DataFrame from flattened results
        if flattened_results:
            result_df = pd.DataFrame(flattened_results)
            
            # Sort by Model, then by Multivalues for better organization
            result_df = result_df.sort_values(['Model', 'Multivalues'])
            
            print(f"Created result DataFrame with {len(result_df)} rows")
            print("\nSample of results:")
            print(result_df.head(10).to_string(index=False))
            
            try:
                # Try to save to the same file with a new sheet
                with pd.ExcelWriter(De_Para_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    # Write the new data
                    result_df.to_excel(writer, sheet_name='tb_de_para', index=False)
                    print(f"Results saved to '{De_Para_file_path}' in sheet 'tb_de_para'")
            except Exception as e:
                # If there's an error (like file is open), save to a new file
                output_path = os.path.splitext(De_Para_file_path)[0] + "_tb_de_para.xlsx"
                result_df.to_excel(output_path, sheet_name='tb_de_para', index=False)
                print(f"Could not update original file: {e}")
                print(f"Results saved to: {output_path}")
            
            return result_df
        else:
            print("No matching results found.")
            return pd.DataFrame()
    
    except Exception as e:
        print(f"Error updating De_Para file: {e}")
        return pd.DataFrame()

# === Resource path function ===
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller .exe"""
    try:
        base_path = sys._MEIPASS  # Set by PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# === GUI Section ===
root = tk.Tk()
root.title("Griglia Transformer")

# Load Stellantis logo image and display it
stellantis_logo_path = resource_path("assets/Vlc_img.png")
try:
    img_stellantis_logo = Image.open(stellantis_logo_path)
    img_stellantis_logo = img_stellantis_logo.resize((530, 40), Image.Resampling.LANCZOS)
    photo_img_stellantis_logo = ImageTk.PhotoImage(img_stellantis_logo)
    root.image = photo_img_stellantis_logo  # Keep reference

    image_frame = tk.Frame(root, bg="#f0f0f0")
    image_frame.pack(pady=10)
    tk.Label(image_frame, image=photo_img_stellantis_logo, bg="#f0f0f0").pack(side="left", padx=20)
except Exception as e:
    print(f"Could not load logo: {e}")

# Input folder selection
tk.Label(root, text="Select Input Folder:").pack(anchor="w", padx=10)
input_folder_var = tk.StringVar()
tk.Entry(root, textvariable=input_folder_var, width=70).pack(padx=10, pady=2)
tk.Button(root, text="Browse Input Folder", command=lambda: input_folder_var.set(filedialog.askdirectory())).pack(pady=5)

# De_Para file selection
tk.Label(root, text="Select De_Para File (optional):").pack(anchor="w", padx=10)
de_para_file_var = tk.StringVar()
tk.Entry(root, textvariable=de_para_file_var, width=70).pack(padx=10, pady=2)
tk.Button(root, text="Browse De_Para File", 
          command=lambda: de_para_file_var.set(filedialog.askopenfilename(
              title="Select De_Para File",
              filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
          ))).pack(pady=5)

# Output folder selection
tk.Label(root, text="Select Output Folder:").pack(anchor="w", padx=10)
output_folder_var = tk.StringVar()
tk.Entry(root, textvariable=output_folder_var, width=70).pack(padx=10, pady=2)
tk.Button(root, text="Browse Output Folder", command=lambda: output_folder_var.set(filedialog.askdirectory())).pack(pady=5)

# Progress bar and label
progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=10)
progress_label = tk.Label(root, text="")
progress_label.pack()

def threaded_processing(input_folder, de_para_file, output_folder):
    try:
        progress_bar.start(10)  # Start progress animation
        progress_label.config(text="Processing...")
        
        output_file_path = os.path.join(output_folder, "Expanded_Mapped_File.xlsx")
        
        transform_and_expand(input_folder, de_para_file, output_file_path)
        
        progress_label.config(text="Done!")
        messagebox.showinfo("Success", f"File saved at:\n{output_file_path}")
        progress_bar.stop()
    except Exception as e:
        messagebox.showerror("Processing Error", str(e))
        print(f"Processing error: {e}")
    finally:
        progress_bar.stop()
        progress_label.config(text="")
        run_button.config(state="normal")

def run_processing():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    de_para_file = de_para_file_var.get()
    
    if not input_folder or not output_folder:
        messagebox.showerror("Error", "Please select both input and output folders.")
        return
    
    # Validate that input folder exists and contains Excel files
    if not os.path.exists(input_folder):
        messagebox.showerror("Error", "Input folder does not exist.")
        return
    
    # Check if there are any Excel files in the input folder
    search_pattern = os.path.join(input_folder, '*.xlsx')
    xlsx_files = glob.glob(search_pattern)
    if not xlsx_files:
        messagebox.showerror("Error", "No Excel files found in the input folder.")
        return
    
    # Validate De_Para file if provided
    if de_para_file and not os.path.exists(de_para_file):
        messagebox.showerror("Error", "Selected De_Para file does not exist.")
        return
    
    # Validate that De_Para file has the required sheet if provided
    if de_para_file:
        try:
            wb = pd.ExcelFile(de_para_file)
            if "Coded" not in wb.sheet_names:
                messagebox.showerror("Error", "De_Para file must contain a 'Coded' sheet.")
                return
        except Exception as e:
            messagebox.showerror("Error", f"Error reading De_Para file: {e}")
            return
    
    # Create output directory if it doesn't exist
    try:
        os.makedirs(output_folder, exist_ok=True)
    except Exception as e:
        messagebox.showerror("Error", f"Cannot create output folder: {e}")
        return
    
    run_button.config(state="disabled")  # Disable button to prevent re-click
    threading.Thread(target=threaded_processing, args=(input_folder, de_para_file, output_folder), daemon=True).start()

run_button = tk.Button(root, text="Run Processing", command=run_processing, height=2, width=30, 
                      bg="#4CAF50", fg="white", font=("Helvetica", 12, "bold"))
run_button.pack(pady=20)

# Add status information
status_frame = tk.Frame(root)
status_frame.pack(pady=10)

tk.Label(status_frame, text="Status Information:", font=("Helvetica", 10, "bold")).pack(anchor="w")
status_text = tk.Text(status_frame, height=8, width=80, wrap=tk.WORD)
status_text.pack(padx=10, pady=5)

# Redirect print statements to the status text widget
class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, str):
        self.widget.insert(tk.END, str)
        self.widget.see(tk.END)
        root.update_idletasks()

    def flush(self):
        pass

# Redirect stdout to the text widget
sys.stdout = TextRedirector(status_text)

# Add a clear button for the status text
def clear_status():
    status_text.delete(1.0, tk.END)

tk.Button(status_frame, text="Clear Status", command=clear_status).pack(pady=5)

# Add instructions
instructions_frame = tk.Frame(root)
instructions_frame.pack(pady=10, padx=10, fill="x")

instructions_text = """
Instructions:
1. Select the input folder containing Excel files with 'Griglia Mondo - Volumi' or 'Grid World - Volume' sheets
2. Optionally select a De_Para file (must contain a 'Coded' sheet with 'Griglia Italiano' and 'Griglia Inglês' columns)
3. Select an output folder where the processed files will be saved
4. Click 'Run Processing' to start the transformation
5. The process will create an 'Expanded_Mapped_File.xlsx' and update the De_Para file if provided

Note: The model code will be extracted from the first 3 characters of each Excel filename
Example: '2261-2025-eGrip TORO MY26 v7 17Fev2025 (12).xlsx' → Model: '226'

The tb_de_para sheet will contain:
- Multivalues: The corresponding values from the Griglia data
- Griglia Italiano: Italian terms from the De_Para file
- Griglia Inglês: English terms from the De_Para file  
- Model: The 3-character code extracted from filename
- Resp.1: Multivalues when matched via Griglia Italiano
- Resp.2: Multivalues when matched via Griglia Inglês
"""

tk.Label(instructions_frame, text=instructions_text, justify="left", wraplength=600, 
         font=("Helvetica", 9)).pack(anchor="w")

# Center the window
root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f"{width}x{height}+{x}+{y}")

root.mainloop()
