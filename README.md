# Computer Class - Student Enrollment Analysis System

A comprehensive data analysis and visualization system for managing and analyzing student enrollment data from computer training classes. This project processes Excel files containing student information and generates interactive dashboards for enrollment analysis.

## 🎯 Project Overview

This system is designed to handle student enrollment data from computer training programs, specifically focusing on:
- **Data Standardization**: Cleans and standardizes student data from multiple Excel files
- **Enrollment Analysis**: Analyzes student enrollment patterns across different years and courses
- **Interactive Dashboards**: Generates visual charts and reports for data insights
- **Data Quality**: Ensures data integrity and identifies duplicate records

## 📊 Features

### Data Processing
- **Multi-file Support**: Processes multiple Excel files simultaneously
- **Column Standardization**: Automatically standardizes column names across different file formats
- **Data Cleaning**: Cleans and normalizes course names, durations, and other fields
- **Gender Extraction**: Automatically extracts gender information from relationship fields
- **Duplicate Detection**: Identifies and reports duplicate student records

### Analytics & Visualization
- **Yearly Enrollment Trends**: Tracks student enrollment across academic years
- **Course-wise Analysis**: Analyzes enrollment patterns by course type
- **Duration Analysis**: Examines course popularity by duration
- **Interactive Charts**: Creates responsive, interactive visualizations using Plotly

### Supported Course Types
- Basic Skills
- Basic + Tally
- Tally
- DIT (Diploma in Information Technology)
- DTP (Desktop Publishing)
- NIIT

## 🛠️ Installation

### Prerequisites
- Python 3.11 or higher
- pip or uv package manager

### Setup Instructions

1. **Clone or download the project files**

2. **Install dependencies using pip:**
   ```bash
   pip install -r requirements.txt
   ```

   **Or using uv (recommended):**
   ```bash
   uv sync
   ```

3. **Verify installation:**
   ```bash
   python main.py
   ```

## 📁 Project Structure

```
Computer_Class/
├── main.py                          # Main entry point
├── extract_excel_data.py            # Data extraction and processing
├── generate_charts.py               # Chart generation and visualization
├── student_enrollment_dashboard.html # Generated interactive dashboard
├── requirements.txt                 # Python dependencies
├── pyproject.toml                   # Project configuration
├── README.md                        # This file
└── Excel Files/
    ├── Updated (07 -10 -2024) Batch 2021 to 2024.xlsx
    └── EDITED NIIT 10 sept 2024.xlsx
```

## 🚀 Usage

### Basic Usage

1. **Run the main analysis:**
   ```bash
   python generate_charts.py
   ```

2. **View the generated dashboard:**
   - Open `student_enrollment_dashboard.html` in your web browser
   - The dashboard contains three interactive charts:
     - Student Enrollment by Year
     - Student Enrollment by Year and Course
     - Course Popularity by Duration

### Data Processing

The system automatically processes Excel files with the following features:

- **Column Mapping**: Standardizes column names like "ADM NO" → "ADM_NO"
- **Year Extraction**: Extracts academic years from sheet names
- **Data Validation**: Checks for data completeness and uniqueness
- **Gender Classification**: Infers gender from relationship fields (S/O, D/O, W/O, H/O)

### Expected Excel Format

Your Excel files should contain sheets with student data including columns like:
- Admission Number
- Student Name
- Father/Husband Name
- Course
- Duration
- Address
- Mobile
- Email
- Education
- Current Status
- Monthly Income

## 📈 Dashboard Features

The generated dashboard provides:

1. **Overall Enrollment Trends**: Bar chart showing total student enrollment by academic year
2. **Course-wise Analysis**: Stacked bar chart showing enrollment by course and year
3. **Duration Analysis**: Bar chart showing course popularity by duration
4. **Interactive Elements**: Hover tooltips, zoom, pan, and legend filtering
5. **Responsive Design**: Adapts to different screen sizes

## 🔧 Configuration

### Customizing Data Sources

To process different Excel files, modify the `excel_files` list in `extract_excel_data.py`:

```python
excel_files = ["your_file_1.xlsx", "your_file_2.xlsx"]
```

### Adding New Course Types

To support additional course types, update the `course_mapping` dictionary in `extract_excel_data.py`:

```python
course_mapping = {
    'YOUR_COURSE': 'Your Course',
    # ... existing mappings
}
```

## 📋 Dependencies

- **pandas** (≥2.0.0): Data manipulation and analysis
- **openpyxl** (≥3.0.0): Excel file reading and writing
- **plotly** (≥5.0.0): Interactive data visualization

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

## 🆘 Support

For issues, questions, or feature requests:
1. Check the existing issues
2. Create a new issue with detailed information
3. Include sample data if reporting bugs

## 🔄 Version History

- **v0.1.0**: Initial release with basic data processing and visualization capabilities

---

**Note**: This system is designed for educational institutions managing computer training programs. Ensure compliance with data protection regulations when processing student information.
