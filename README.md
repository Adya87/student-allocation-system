# 📊 Student Activity Allocation System

## 📌 Overview
This project automates the allocation of student activities based on their preferences and predefined capacity limits. It also includes reconciliation of student data across academic years to track submissions and identify missing or withdrawn students.

Designed to simplify administrative workflows in schools using automation.

## 🚀 Features
- Allocates activities based on student preferences
- Handles capacity constraints per grade
- Removes duplicate student entries
- Supports different input formats (Grade 8 & others)
- Generates separate Excel outputs per grade
- Reconciles previous and next grade student data
- Highlights:
  - Students who did not submit forms
  - Students who left school (in red)

## 🛠️ Tech Stack
- Python
- Pandas
- OpenPyXL

## 📂 Project Structure
```
.
├── allocation.py        # Activity allocation logic
├── reconciliation.py    # Data reconciliation logic
├── sample_data/         # Example Excel files (optional)
└── README.md
```

## ▶️ How to Run

### 1. Install dependencies
```bash
pip install pandas openpyxl
```

### 2. Run allocation
```bash
python allocation.py
```
Then enter the input file path when prompted.

### 3. Run reconciliation
```bash
python reconciliation.py
```
Then enter the required file paths when prompted.

## 📊 Output
- Separate Excel files for each grade (allocation)
- Final reconciled Excel file with:
  - Submission status
  - Left school students
  - Red-highlighted rows for special cases

## ⚠️ Notes
- Ensure Excel column formats match expected structure
- Names are normalized using first and last name only
- Input file paths are entered during runtime

## 👤 Author
Adya Tyagi
