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
   - The dashboard contains eight interactive charts and visualizations

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

## 📈 Comprehensive Dashboard Charts

The system generates a comprehensive dashboard with **8 interactive charts** that provide deep insights into student enrollment and outcomes:

### 1. 📊 Student Enrollment by Year
- **Chart Type**: Bar Chart
- **Purpose**: Shows overall student enrollment trends across academic years
- **Features**: 
  - Displays total number of unique students per year
  - Interactive hover tooltips with exact counts
  - Chronological year ordering
  - Sky blue color scheme for easy reading

### 2. 📈 Student Enrollment by Year and Course
- **Chart Type**: Grouped Bar Chart
- **Purpose**: Analyzes enrollment patterns by both year and course type
- **Features**:
  - Color-coded bars for each course type
  - Side-by-side comparison across years
  - Custom legend showing all course types
  - Hover tooltips with year, course, and student count

### 3. 🎯 Course Popularity by Duration
- **Chart Type**: Horizontal Bar Chart
- **Purpose**: Shows which course-duration combinations are most popular
- **Features**:
  - Sorted by popularity (highest to lowest)
  - Combined course and duration labels
  - Color-coded by course type
  - 45-degree rotated labels for better readability

### 4. 👥 Gender Distribution (Overall)
- **Chart Type**: Donut Chart
- **Purpose**: Displays the overall gender distribution of all students
- **Features**:
  - Percentage and count labels
  - Blue for Male, Pink for Female
  - Excludes "Unknown" gender classifications
  - Interactive hover with detailed statistics

### 5. 📊 Gender Distribution per Year
- **Chart Type**: Grouped Bar Chart
- **Purpose**: Shows gender distribution trends over time
- **Features**:
  - Year-wise breakdown of male/female enrollment
  - Consistent color coding with overall gender chart
  - Chronological year ordering
  - Hover tooltips with year, gender, and count

### 6. 💼 Total vs Employed Students per Year
- **Chart Type**: Grouped Bar Chart
- **Purpose**: Compares total enrollment with employment outcomes
- **Features**:
  - Light blue bars for total students
  - Green bars for employed students
  - Employment rate visualization
  - Year-wise employment tracking

### 7. 🎓 Employed Students by Year and Course
- **Chart Type**: Grouped Bar Chart
- **Purpose**: Shows employment outcomes by course type over time
- **Features**:
  - Course-specific employment tracking
  - Color-coded by course type
  - Year-wise employment patterns
  - Helps identify most employable courses

### 8. 🏆 Student Outcomes After Course Completion
- **Chart Type**: Pie Chart
- **Purpose**: Shows what students are doing after completing their courses
- **Features**:
  - Categories: Employed, Student, Homemaker, Jobseeker, Other
  - Color-coded status categories
  - Percentage and count labels
  - Overall outcome assessment

## 🎨 Interactive Features

### Dashboard Capabilities
- **Responsive Design**: Adapts to different screen sizes
- **Interactive Elements**: 
  - Hover tooltips with detailed information
  - Zoom and pan functionality
  - Legend filtering
  - Click-to-highlight features
- **Custom Legends**: Positioned on the right side for easy reference
- **Professional Styling**: Clean, modern design with consistent color schemes

### Data Quality Features
- **Uniqueness Validation**: Checks for duplicate student records
- **Data Completeness**: Identifies missing or incomplete data
- **Statistical Summary**: Provides enrollment and employment statistics
- **Error Reporting**: Highlights data quality issues

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

### Chart Customization

The dashboard can be customized by modifying `generate_charts.py`:
- **Colors**: Update color maps for different categories
- **Layout**: Adjust chart positioning and sizing
- **Interactivity**: Modify hover templates and click behaviors
- **Styling**: Change fonts, borders, and overall appearance

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

- **v0.1.0**: Initial release with comprehensive 8-chart dashboard and data processing capabilities

---

**Note**: This system is designed for educational institutions managing computer training programs. Ensure compliance with data protection regulations when processing student information.
