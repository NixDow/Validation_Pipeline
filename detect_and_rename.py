from rapidfuzz import process
from config import HEADER_SCAN_DEPTH, COLUMN_MATCH_THRESHOLD

def detect_header_row(ws, expected_columns):

    best_row = None
    best_score = 0

    for row in range(1, HEADER_SCAN_DEPTH + 1): #iterate through the rows of potential headers until the HEADER_SCAN_DEPTH limit

        cells = [str(cell.value) for cell in ws[row] if cell.value] #extract the text from each row in the worksheet

        score = 0

        for expected in expected_columns:

            match = process.extractOne(expected, cells) #compares the extracted cell to that of the expected columns

            if match and match[1] > 70: #Each time a cell matches with one in the workflow with a score > 70 its overall score increases
                score += 1

        if score > best_score:
            best_score = score
            best_row = row

    return best_row #returns the index of the row with the highest score - our likely row headers

def map_columns(expected_columns, actual_columns):

    mapping = {}
    actual_columns = [str(c) for c in actual_columns if c is not None]

    for expected in expected_columns: #goes through expected columns for each worksheet

        match = process.extractOne(expected, actual_columns)  #attemps to find the highest matching cell in the headers row

        if match and match[1] >= COLUMN_MATCH_THRESHOLD: #matching threshold

            mapping[expected] = match[0]  #Returns a key value pair in the dictionary (expected column: column in worksheet)

    return mapping

def get_max_row(ws):
    """
    Finds the last row that actually contains data,
    preventing validation of 'ghost' rows.
    """
    # Iterate from the very bottom of the sheet upwards
    for row in range(ws.max_row, 0, -1):
        # Check if any cell in this row has a value
        if any(cell.value is not None for cell in ws[row]):
            return row
    return 0
