from concurrent.futures import ProcessPoolExecutor, as_completed
from config import INPUT_FOLDER
from file_processing import validate_file

MAX_WORKERS = 5

def process_file(file_path):
    try:
        print(f"Processing {file_path.name}")
        issues = validate_file(file_path)
        return (file_path.name, f"SUCCESS - {issues} issues flagged")
    except Exception as e:
        return (file_path.name, f"Error: {str(e)}")

def run():
    folder = INPUT_FOLDER

    # 1. Explicitly fetch both extensions and combine them into one list
    all_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm"))

    # 2. Filter out temporary owner files (those starting with ~$)
    files = [f for f in all_files if not f.name.startswith("~$")]

    results = []

    if not files:
        print(f"No .xlsx or .xlsm files found in {folder}. Run generate_messy_data.py first.")
        return

    with ProcessPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [
            executor.submit(process_file, file)
            for file in files
        ]

        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            print(result)

    print("\nSummary:")
    for r in results:
        print(r)

if __name__ == "__main__":
    run()
