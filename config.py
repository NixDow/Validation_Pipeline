#schema for each tab and validation rules for each field
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
INPUT_FOLDER = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output" #validated (labelled + reformatted) copies of the input files
REPORT_OUTPUT_DIR = BASE_DIR / "reports"

HEADER_SCAN_DEPTH = 8
COLUMN_MATCH_THRESHOLD = 90 #fuzzy matching threshold
BLANK_FILL = "FFEA00" #bright yellow
INVALID_FILL = "50C878" #green
MISMATCH_FILL = "00FFFF" #light blue

#plausible range for any date in the submission
MIN_DATE = "2020/01/01"
MAX_DATE = None #None = today

QTY_LIMIT = 50 #quantities above this are flagged (same idea as recruit_limit)

DROP_DOWN_ENTRIES = {
    "Stores": ["Camden Cafe",
               "Brixton Cafe",
               "Shoreditch Cafe",
               "Richmond Cafe"
            ],
    "Category": ["Drink", "Food"],
    "Item": ["Coffee",
             "Tea",
             "Juice",
             "Smoothie",
             "Sandwich",
             "Salad",
             "Cake",
             "Cookie"
            ],
    "Payment": ["Cash", "Card", "Digital Wallet"],
    "Status": ["Completed", "Refunded", "Cancelled", "Pending"],
}

#known variants that get rewritten to the correct dropdown value before the dropdown check

STANDARDISE_VALUES = {
    "Category": {
        "drinks": "Drink",
        "beverage": "Drink",
        "foods": "Food",
    },
    "Payment": {
        "credit card": "Card",
        "debit card": "Card",
        "cc": "Card",
        "e-wallet": "Digital Wallet",
        "digital-wallet": "Digital Wallet",
        "apple pay": "Digital Wallet",
    },
    "Status": {
        "complete": "Completed",
        "done": "Completed",
        "canceled": "Cancelled",
        "refund": "Refunded",
    },
}

VALIDATION_RULES = {

    "Store Details": {
        "expected_columns": {}
    },
    "Orders": {
        "expected_columns": {
            "Order ID": ["not_blank"],
            "Order Date": ["not_blank", "date", "date_range"],
            "Item Category": ["not_blank", "list:Category"],
            "Item": ["not_blank", "list:Item"],
            "Quantity": ["number", "qty_limit"], #conditional validation
            "Unit Price": ["not_blank", "number"],
            "Payment Method": ["not_blank", "list:Payment"],
            "Order Status": ["not_blank", "list:Status"],
            "Refund Date": ["date", "date_range"], #conditional validation
            "Customer Email": ["email"],
            "Notes": []
        }
    },
    "conditional_rules": [
        {"type": "required_if",
         "if_column": "Order Status",
         "if_value": "Refunded",
         "then_column": "Refund Date"
        },
        {"type": "required_if_list",
         "if_column": "Order Status",
         "if_values": ["Completed", "Refunded"],
         "then_column": "Quantity"
        },
        {"type": "blanks_or_zero_list",
         "if_column": "Order Status",
         "if_values": ["Cancelled"],
         "then_column": "Quantity"
        },
        {"type": "dropdown_dependency",
         "source_column": "Item Category",
         "target_column": "Item",
         "mapping": {
             "Drink": ["Coffee", "Tea", "Juice", "Smoothie"],
             "Food": ["Sandwich", "Salad", "Cake", "Cookie"]
            }
        },
        {"type": "date_not_before",
         "if_column": "Order Date",
         "then_column": "Refund Date"
        }
    ]
}
