# Excel Automation Toolkit

A desktop application built with **Python**, **Pandas**, and **Tkinter** to automate repetitive Microsoft Excel workflows commonly encountered during enterprise **data migration** and **data validation** projects.

---

## Background

This project was developed during my internship to support enterprise data migration activities.

One of the most time-consuming stages of a data migration project is **manual data validation** (internally referred to as *data eyeballing*), where thousands of records must be reviewed to ensure data integrity before migration.

The extracted datasets often contained tens of thousands of records, making them difficult to review and distribute among team members. Different validation teams also required different batch sizes depending on their assigned workload.

To reduce manual effort and improve productivity, I developed the **Excel Automation Toolkit**, a desktop application that automates several repetitive Excel processing tasks.

---

## Features

### 📂 Excel Splitter

Split large Excel workbooks into multiple smaller files based on a user-defined number of rows.

Ideal for:

- Batch processing
- Team workload distribution
- System upload limitations

---

### 🎲 Random Excel Extractor

Extract a reproducible random sample from an Excel file.

Features:

- Percentage-based extraction
- Fixed random seed
- Original Excel row reference
- Quality assurance and audit sampling

---

### 🧹 Remove Records From Reference

Compare a source dataset with a reference dataset and automatically remove matching records using a selected primary key.

Generated outputs include:

- Cleaned dataset
- Removed records
- Processing summary report

---

## Why This Project?

Instead of performing repetitive Excel operations manually, this toolkit allows users to complete common data preparation tasks with only a few clicks.

The toolkit helps reduce:

- Manual Excel work
- Human errors
- Processing time
- Repetitive data migration tasks

---

## Technologies Used

- Python
- Pandas
- Tkinter
- OpenPyXL

---

## Requirements

Python 3.8+

```bash
pip install -r requirements.txt
```

---

## Usage

Run the application:

```bash
python excel_toolkit_gui.py
```

Then:

1. Choose the required tool.
2. Select the input file(s).
3. Select the output folder.
4. Configure the required options.
5. Start the process.

---

## Repository Structure

```
excel-automation-toolkit/
│
├── excel_toolkit_gui.py
├── requirements.txt
├── LICENSE
├── README.md
└── legacy/
```

The **legacy** folder contains earlier standalone scripts that were merged into the unified toolkit.

---

## Future Enhancements

- Merge multiple Excel files
- Duplicate record cleaner
- Excel file comparison
- Multi-sheet support
- Drag-and-drop file upload
- Dark mode
- Multi-threaded processing

---

## Packaging (Optional)

Build a standalone Windows executable:

```bash
pip install pyinstaller

pyinstaller --noconsole --onefile --name "Excel Automation Toolkit V1" excel_toolkit_gui.py
```

---

## License

This project is licensed under the MIT License.