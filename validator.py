import re
from datetime import datetime, date
from config import DROP_DOWN_ENTRIES, STANDARDISE_VALUES, MIN_DATE, MAX_DATE

DATE_FORMATS = [
    "%d/%m/%Y", #16/09/2016
    "%d.%m.%Y", #16.09.2016
    "%d-%m-%Y", #16-09-2016
    "%d/%m/%y", #31/06/26
    "%Y/%m/%d", #2016/09/16
    "%Y.%m.%d", #2016.09.16
    "%Y-%m-%d", #2016-09-16
    "%Y-%m-%d %H:%M:%S", #2016-09-16 00:00:00
]

MONTH_NAME_FORMATS = [
    "%d %b %Y",   # 01 Jan 2024
    "%d %B %Y",   # 01 January 2024
    "%b %d %Y",   # Jan 01 2024
    "%B %d, %Y",  # January 01, 2024
]

OUTPUT_DATE_FORMAT = "%Y/%m/%d"

#Checking for blanks function
def is_blank(value):
    """
    Returns True if a value is empty or None
    """
    return value is None or str(value).strip() == ""

#Checks for floats and integers
def make_number(value):
    """
    Attempts to convert a value to a float.
    Returns:
        - The numeric value if successful
        - The term 'invalid' if conversion failed
        - None if the value is blank
    """
    if is_blank(value): #skip empty cells
        return None

    if isinstance(value, bool): #True/False are not valid numbers
        return 'invalid'

    if isinstance(value, (float, int)): # if cell is a number or float, return the actual value
        return value

    val_str = str(value).strip()
    if val_str.startswith('='):
        return value

    try:
        cleaned_val = val_str.replace('£', '').replace(',', '').replace(' ', '')

        return float(cleaned_val) #try to convert string e.g '0' to a number

    except (ValueError, TypeError): #if it can't be converted, return invalid
        return 'invalid'

def standardise_value(value, list_key):
    """
    Rewrites a known variant (e.g. 'credit card') to its dropdown value (e.g. 'Card').
    Also fixes the case of values that match a dropdown entry case-insensitively.
    Returns the value unchanged if no fix is known.
    """
    if is_blank(value):
        return value

    check_val = str(value).strip().lower()

    fixes = STANDARDISE_VALUES.get(list_key, {})
    if check_val in fixes:
        return fixes[check_val]

    for allowed in DROP_DOWN_ENTRIES.get(list_key, []):
        if check_val == allowed.lower():
            return allowed

    return value

def drop_downs_list(value, list_key):
    """
    Checks if a value in a specific column is present within the dropdown list
    """
    if is_blank(value):
        return True

    #get the allowed list from config
    allowed_raw = DROP_DOWN_ENTRIES.get(list_key,[])
    allowed_ins = [str(v).strip().lower() for v in allowed_raw]

    entries_to_check = [item.strip().lower() for item in str(value).split(';')]

    return all(entry in allowed_ins for entry in entries_to_check if entry)

def parse_with_month_names(value):
    """
    Parsing dates that have the month as a name
    """
    for fmt in MONTH_NAME_FORMATS:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue

    return None

def parse_date(value):
    """
    Returns a datetime for any recognised date input, otherwise None
    """
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    if isinstance(value, date): #if a value in spreadsheet is already classed as a date
        return datetime(value.year, value.month, value.day)

    value = str(value).strip()

    # Try parsing using known formats
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue

    return parse_with_month_names(value)

def parse_and_standardise_date(value):
    """
    Check the date format from the value and re-format if the date is valid
    """
    parsed = parse_date(value)
    return parsed.strftime(OUTPUT_DATE_FORMAT) if parsed else None

def val_and_form_date(value):
    """
    Logic:
    - if is_blank is True -> 'blank'
    - If date is valid -> 'valid'
    - Else -> return 'invalid'
    """
    if is_blank(value):
        return "blank", None

    parsed = parse_and_standardise_date(value)
    if parsed:
        return "valid", parsed
    return "invalid", None

def in_date_range(standardised_date):
    """
    Checks a standardised date (YYYY/MM/DD) sits between MIN_DATE and MAX_DATE (default today).
    Catches dates that parse but are unreasonable, e.g. 1999 or next decade.
    """
    parsed = datetime.strptime(standardised_date, OUTPUT_DATE_FORMAT)
    min_date = datetime.strptime(MIN_DATE, OUTPUT_DATE_FORMAT)
    max_date = datetime.strptime(MAX_DATE, OUTPUT_DATE_FORMAT) if MAX_DATE else datetime.now()

    return min_date <= parsed <= max_date

def apply_conditions(ws, row, column_map, rules, headers):
    """
    Apply conditional validation logic using case-insensitive comparisons.
    """
    errors = []

    for rule in rules:
        if_col_name = rule.get("if_column") or rule.get("source_column")
        source_col = column_map.get(if_col_name)

        if not source_col:
            continue

        source_index = headers.index(source_col) + 1
        # Normalize source value for comparison
        raw_source = ws.cell(row=row, column=source_index).value
        source_val_clean = str(raw_source).strip().lower() if raw_source is not None else ""

        # 1. 'required_one_of'
        if rule["type"] == "required_one_of":
            trigger_val = str(rule["if_value"]).strip().lower()
            if source_val_clean == trigger_val:
                valid_found = False
                targets_to_flag = []
                # Normalize exclusion list to uppercase for robust matching
                exclude_list = [str(v).strip().upper() for v in rule.get("exclude_values", [])]

                for col_name in rule["then_columns"]:
                    actual_col = column_map.get(col_name)
                    if actual_col:
                        idx = headers.index(actual_col) + 1
                        val = ws.cell(row=row, column=idx).value
                        if not is_blank(val):
                            if str(val).strip().upper() not in exclude_list:
                                valid_found = True
                                break
                        targets_to_flag.append(col_name)

                if not valid_found:
                    for col_name in targets_to_flag:
                        errors.append((col_name, "required_if"))

        # 2. 'required_if'
        elif rule["type"] == "required_if":
            target_col = column_map.get(rule["then_column"])
            trigger_val = str(rule["if_value"]).strip().lower()
            if target_col and source_val_clean == trigger_val:
                target_val = ws.cell(row=row, column=headers.index(target_col) + 1).value
                if is_blank(target_val):
                    errors.append((rule["then_column"], "required_if"))

        # 3. 'required_if_list'
        elif rule["type"] == "required_if_list":
            target_col = column_map.get(rule["then_column"])
            # Normalize the comparison list to lowercase
            if_list = [str(v).strip().lower() for v in rule["if_values"]]
            if target_col and source_val_clean in if_list:
                target_val = ws.cell(row=row, column=headers.index(target_col) + 1).value
                if is_blank(target_val):
                    errors.append((rule["then_column"], "required_if"))

        # 4. 'blanks_or_zero_list'
        elif rule["type"] == "blanks_or_zero_list":
            target_col = column_map.get(rule["then_column"])
            if_list = [str(v).strip().lower() for v in rule["if_values"]]
            if target_col and source_val_clean in if_list:
                target_val = ws.cell(row=row, column=headers.index(target_col) + 1).value
                is_zero = False
                try:
                    if float(target_val) == 0: is_zero = True
                except (ValueError, TypeError): pass

                if not (is_blank(target_val) or is_zero):
                    errors.append((rule["then_column"], "blanks_or_zero_list"))

        # 5. 'dropdown_dependency'
        elif rule["type"] == "dropdown_dependency":
            target_col = column_map.get(rule["target_column"])
            if target_col:
                target_val = ws.cell(row=row, column=headers.index(target_col) + 1).value
                if is_blank(target_val):
                    continue
                # Create a lowercase-keyed dictionary for case-insensitive lookup
                lower_mapping = {str(k).lower(): v for k, v in rule["mapping"].items()}
                allowed_options = lower_mapping.get(source_val_clean)

                if allowed_options:
                    clean_target = str(target_val).strip().lower()
                    lower_allowed_t = [str(o).strip().lower() for o in allowed_options]

                    if clean_target not in lower_allowed_t:
                        errors.append((rule["target_column"], "dropdown_mismatch"))

        # 6. 'date_not_before' - then_column date must be on or after if_column date
        elif rule["type"] == "date_not_before":
            target_col = column_map.get(rule["then_column"])
            if target_col:
                earlier = parse_date(raw_source)
                later = parse_date(ws.cell(row=row, column=headers.index(target_col) + 1).value)
                if earlier and later and later < earlier:
                    errors.append((rule["then_column"], "date_order"))

    return errors

def valid_email(value):
    """
    Validates the format of an email address
    """
    if is_blank(value):
        return True

    pattern = r"^[\w.+-]+@[\w-]+(\.[\w-]+)*\.[a-zA-Z]{2,}$"

    return bool(re.match(pattern, str(value).strip()))

#Order ID cleaning
def clean_order_id(order_id, notes):
    """
    Cleans and salvages Order IDs into the standard ORD-00000 format.
    - '#ord-12', 'ORD 12', 'ord_0012' -> 'ORD-00012'
    - blank ID -> salvaged from the Notes column if a reference is present
    Returns the original value if it cannot be cleaned.
    """
    if not is_blank(order_id):
        id_str = str(order_id).strip()
        match = re.fullmatch(r'#?\s*ord[\s_\-]*(\d{1,5})', id_str, flags=re.IGNORECASE)
        if match:
            return f"ORD-{int(match.group(1)):05d}"
        return order_id

    # Salvage Order ID from Notes if ID is blank
    if notes:
        match = re.search(r'ORD[\s_\-]*(\d{1,5})', str(notes), flags=re.IGNORECASE)
        if match:
            return f"ORD-{int(match.group(1)):05d}"

    return order_id
