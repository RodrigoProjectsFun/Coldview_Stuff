import polars as pl
import glob
import os
import concurrent.futures

def fast_polars_workers(source_folder, target_file):
    print("Loading target cards...")
    df_targets = pl.read_excel(target_file)
    
    files = glob.glob(os.path.join(source_folder, "*.xlsx"))
    if not files:
        print("No files found.")
        return

    # --- THE WORKER FUNCTION ---
    def read_single_excel(file_path):
        try:
            # We strictly use the Rust-based calamine engine for speed
            df = pl.read_excel(
                file_path, 
                engine="calamine", 
                read_options={"columns": ["Card Number", "Account Number"]}
            )
            print(f"Loaded: {os.path.basename(file_path)}")
            
            # Immediately convert the eager DataFrame into a lazy one for the final merge
            return df.lazy()
            
        except Exception as e:
            print(f"Skipping {os.path.basename(file_path)}: {e}")
            return None

    print(f"\nSpinning up thread workers to read {len(files)} files...")
    
    # --- THREAD POOL EXECUTION ---
    # We use ThreadPoolExecutor instead of ProcessPoolExecutor. 
    # Threads share memory, avoiding the expensive serialization of Polars dataframes.
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = list(executor.map(read_single_excel, files))

    # Filter out any failed reads
    lazy_frames = [df for df in results if df is not None]

    if not lazy_frames:
        print("No valid data to process.")
        return

    print("\nConsolidating and merging data (Polars native multi-threading takes over here)...")
    
    # 1. Concat all the lazy frames into one master lookup
    # 2. Drop duplicates
    master_lookup = pl.concat(lazy_frames).unique(subset=["Card Number"], keep="last")
    
    # 3. Perform the Left Join
    # 4. .collect() tells Polars to execute the entire optimized query using all CPU cores at once
    df_final = df_targets.lazy().join(master_lookup, on="Card Number", how="left").collect()
    
    # Save the result
    output_path = "polars_threaded_output.xlsx"
    df_final.write_excel(output_path)
    print(f"\nSuccess! Saved to {output_path}")

if __name__ == "__main__":
    fast_polars_workers(
        source_folder=r'C:\path\to\your\source\excels',
        target_file=r'C:\path\to\cards_to_search.xlsx'
    )