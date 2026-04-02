import polars as pl
import glob
import os
import concurrent.futures
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class DataMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Reconciliation Data Matcher")
        self.root.geometry("550x380")
        self.root.configure(padx=20, pady=20)
        
        # --- Variables ---
        self.source_folder = tk.StringVar()
        self.target_file = tk.StringVar()
        self.search_col = tk.StringVar(value="Card Number")
        self.return_col = tk.StringVar(value="Account Number")
        
        self.build_ui()

    def build_ui(self):
        # --- Source Folder Selection ---
        tk.Label(self.root, text="1. Select Source Folder (Contains all daily extracts):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        frame_source = tk.Frame(self.root)
        frame_source.pack(fill="x", pady=(0, 15))
        tk.Entry(frame_source, textvariable=self.source_folder, state="readonly", width=50).pack(side="left", padx=(0, 10))
        tk.Button(frame_source, text="Browse...", command=self.browse_folder).pack(side="left")

        # --- Target File Selection ---
        tk.Label(self.root, text="2. Select Target File (The list of cards to search):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        frame_target = tk.Frame(self.root)
        frame_target.pack(fill="x", pady=(0, 15))
        tk.Entry(frame_target, textvariable=self.target_file, state="readonly", width=50).pack(side="left", padx=(0, 10))
        tk.Button(frame_target, text="Browse...", command=self.browse_file).pack(side="left")

        # --- Column Configuration ---
        tk.Label(self.root, text="3. Configure Columns:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        frame_cols = tk.Frame(self.root)
        frame_cols.pack(fill="x", pady=(0, 20))
        
        tk.Label(frame_cols, text="Search Column:").pack(side="left")
        tk.Entry(frame_cols, textvariable=self.search_col, width=18).pack(side="left", padx=(5, 15))
        
        tk.Label(frame_cols, text="Return Column:").pack(side="left")
        tk.Entry(frame_cols, textvariable=self.return_col, width=18).pack(side="left", padx=(5, 0))

        # --- Action Button & Status ---
        self.run_btn = tk.Button(self.root, text="▶ RUN MATCHING PROCESS", font=("Segoe UI", 11, "bold"), bg="#4CAF50", fg="white", command=self.start_processing)
        self.run_btn.pack(fill="x", pady=10, ipady=5)
        
        self.status_label = tk.Label(self.root, text="Ready.", fg="gray", font=("Segoe UI", 9))
        self.status_label.pack()

    # --- UI Interactions ---
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.source_folder.set(folder)

    def browse_file(self):
        file = filedialog.askopenfilename(title="Select Target File", filetypes=[("Excel Files", "*.xlsx")])
        if file:
            self.target_file.set(file)

    def update_status(self, message, color="black"):
        self.status_label.config(text=message, fg=color)
        self.root.update_idletasks()

    # --- Execution Logic ---
    def start_processing(self):
        # Validate inputs
        if not self.source_folder.get() or not self.target_file.get():
            messagebox.showwarning("Missing Input", "Please select both the source folder and the target file.")
            return
        if not self.search_col.get() or not self.return_col.get():
            messagebox.showwarning("Missing Input", "Please specify both column names.")
            return

        # Disable button and update UI
        self.run_btn.config(state="disabled", text="PROCESSING... PLEASE WAIT", bg="#81C784")
        self.update_status("Starting background threads...", "blue")
        
        # Run the heavy Polars logic in a background thread so the GUI doesn't freeze
        threading.Thread(target=self.run_polars_logic, daemon=True).start()

    def run_polars_logic(self):
        source = self.source_folder.get()
        target = self.target_file.get()
        search_c = self.search_col.get().strip()
        return_c = self.return_col.get().strip()
        
        try:
            self.update_status("Loading target cards...")
            df_targets = pl.read_excel(target)
            
            files = glob.glob(os.path.join(source, "*.xlsx"))
            if not files:
                self.update_status("Error: No Excel files found in the source folder.", "red")
                self.reset_button()
                return

            self.update_status(f"Reading {len(files)} files via concurrent threads...")
            
            # The Worker
            def read_single_excel(file_path):
                try:
                    df = pl.read_excel(
                        file_path, 
                        engine="calamine", 
                        read_options={"columns": [search_c, return_c]}
                    )
                    return df.lazy()
                except Exception:
                    return None

            # Thread Pool execution
            with concurrent.futures.ThreadPoolExecutor() as executor:
                results = list(executor.map(read_single_excel, files))

            lazy_frames = [df for df in results if df is not None]

            if not lazy_frames:
                self.update_status("Error: No valid data found. Check your column names.", "red")
                self.reset_button()
                return

            self.update_status("Consolidating and executing Polars high-speed merge...")
            
            # Native Polars magic
            master_lookup = pl.concat(lazy_frames).unique(subset=[search_c], keep="last")
            df_final = df_targets.lazy().join(master_lookup, on=search_c, how="left").collect()
            
            # Save the result in the same directory as the target file
            output_dir = os.path.dirname(target)
            output_path = os.path.join(output_dir, "MATCHED_RESULTS.xlsx")
            
            self.update_status("Saving final output file...")
            df_final.write_excel(output_path)
            
            self.update_status("Success!", "green")
            messagebox.showinfo("Process Complete", f"Data matched successfully!\n\nResults saved to:\n{output_path}")

        except Exception as e:
            self.update_status("An error occurred.", "red")
            messagebox.showerror("Execution Error", f"An error occurred during processing:\n\n{str(e)}")
            
        finally:
            self.reset_button()

    def reset_button(self):
        self.run_btn.config(state="normal", text="▶ RUN MATCHING PROCESS", bg="#4CAF50")

if __name__ == "__main__":
    # Create the main window and start the application
    root = tk.Tk()
    app = DataMatcherApp(root)
    root.mainloop()