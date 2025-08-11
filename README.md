# Excel School Data Merger Web App

A web application that merges duplicate school records in Excel files based on School Number.

## Features

- 📊 **Smart Data Merging**: Groups records by School Number and intelligently merges data
- 💰 **Financial Aggregation**: Automatically sums revenue and student count fields
- 📝 **Text Combination**: Merges text fields intelligently (keeps unique values)
- ⏰ **Timestamped Output**: Generates unique filenames to prevent overwrites
- 🌐 **Web Interface**: Easy-to-use drag-and-drop file upload
- 🎨 **Modern UI**: Beautiful, responsive design

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the Application

```bash
python app.py
```

### 3. Open in Browser

Visit: http://localhost:5000

### 4. Upload Your Excel File

- Drag and drop or click to select your .xlsx file
- Click "Process & Download Merged File"
- Your processed file will automatically download

## File Requirements

Your Excel file should have:
- ✅ `.xlsx` format
- ✅ A column named "School No" (used for grouping)
- ✅ Headers in the first row
- ✅ Financial columns you want to aggregate

## Fields That Get Summed

The following columns are automatically summed when merging duplicate schools:

- Total Order Value (Exclusive GST)
- Total Order Value (Inclusive GST)
- ASSET Revenue
- ASSETStudents
- CARES Revenue
- CARESStudents
- Mindspark Revenue
- MindsparkStudents

## How It Works

1. **Upload**: Select your Excel file through the web interface
2. **Process**: The app groups all rows by "School No"
3. **Merge**: For each school:
   - Financial fields are summed
   - Text fields are merged (same values kept, different values combined)
   - School numbers are preserved
4. **Download**: Get your processed file with timestamp

## Project Structure

```
excel-merger/
├── app.py                 # Main Flask application
├── merger.py              # Excel processing logic
├── requirements.txt       # Python dependencies
├── templates/
│   ├── index.html        # Upload page
│   └── about.html        # About page
├── uploads/              # Temporary uploaded files
└── outputs/              # Processed output files
```

## Example

**Before (Multiple rows for same school):**
```
School No | School Name | ASSET Revenue | CARES Revenue
123       | ABC School  | 10000        | 5000
123       | ABC School  | 15000        | 3000
```

**After (Merged into single row):**
```
School No | School Name | ASSET Revenue | CARES Revenue
123       | ABC School  | 25000        | 8000
```

## Technical Details

- **Backend**: Python Flask
- **Data Processing**: pandas library
- **File Handling**: openpyxl for Excel files
- **Frontend**: HTML5, CSS3, JavaScript (drag-and-drop)

## Security Features

- File type validation (only .xlsx allowed)
- Secure filename handling
- Temporary file cleanup
- Input sanitization

## Customization

To modify which fields get summed, edit the `SUM_COLUMNS` list in `merger.py`:

```python
SUM_COLUMNS = [
    "Your Custom Field 1",
    "Your Custom Field 2",
    # Add more fields as needed
]
```

## Troubleshooting

**"No module named 'flask'"**: Run `pip install -r requirements.txt`

**File upload fails**: Ensure your file is in .xlsx format and has a "School No" column

**Processing error**: Check that your Excel file has proper headers and data format

## License

This project is open source and available under the MIT License.
