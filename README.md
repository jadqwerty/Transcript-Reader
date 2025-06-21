# Transcript Reader

A GUI Windows executable application to extract and convert academic transcripts in PDF format to structured Excel files.

The application supports:
- Extracting student names, IDs, majors, semesters, and course details from transcript PDFs.
- Merging multi-page PDFs correctly into a single Excel file.
- Combining all generated Excel files into a single summary Excel file.
- Real-time log display inside the application.
- Popup error messages for processing issues.
- Custom application icon and resource handling.

---

## 🚀 How to Run

Simply double-click the provided `app.exe` file inside the project folder.  
No Python installation is required.

---

## 📂 Folder Structure
```text
Transcript-REader/
│
├── app.py             # The source code (optional, for modification, in terminal run: python -m PyInstaller --onefile --windowed --icon=icon.ico --add-data "icon.png;." app.py) 
├── app.exe            # The executable to run
├── icon.ico           # The app icon
├── icon.png           # Image resource
└── README.md
