import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import sys
import os

# Import the data processing functions from extract_excel_data.py
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from extract_excel_data import sheet_data, file_sheet_dict

def check_combination_uniqueness(df):
    """
    Check if the combination of ADM_NO, STUDENT, FATHER_HUSBAND, COURSE, DURATION, YEAR is unique
    """
    # Define the columns to check for uniqueness
    columns_to_check = ['ADM_NO', 'STUDENT', 'FATHER_HUSBAND', 'COURSE', 'DURATION', 'YEAR']
    
    # Check if all required columns exist
    missing_columns = [col for col in columns_to_check if col not in df.columns]
    if missing_columns:
        print(f"Warning: Missing columns: {missing_columns}")
        return None
    
    # Create the combination
    combination = df[columns_to_check].astype(str)
    
    # Check for duplicates
    duplicates = combination.duplicated(keep=False)
    duplicate_count = duplicates.sum()
    total_records = len(df)
    
    print(f"\n=== Uniqueness Check Results ===")
    print(f"Total records: {total_records}")
    print(f"Duplicate combinations: {duplicate_count}")
    print(f"Unique combinations: {total_records - duplicate_count}")
    print(f"Uniqueness percentage: {((total_records - duplicate_count) / total_records * 100):.2f}%")
    
    if duplicate_count > 0:
        print(f"\n=== Duplicate Records Found ===")
        duplicate_records = df[duplicates].sort_values(columns_to_check)
        print(f"First few duplicate records:")
        print(duplicate_records[columns_to_check].head(10))
        
        # Show duplicate combinations with their counts
        duplicate_combinations = combination[duplicates].value_counts()
        print(f"\nDuplicate combinations and their counts:")
        print(duplicate_combinations.head(10))
    else:
        print("✓ All combinations are unique!")
    
    return duplicate_count == 0

def create_combined_charts(df):
    """
    Create a single HTML file with both charts
    """
    # Create subplots with 3 rows and 1 column
    fig = make_subplots(
        rows=3, cols=1,
        subplot_titles=['Student Enrollment by Year', 'Student Enrollment by Year and Course', 'Course Popularity by Duration'],
        vertical_spacing=0.08,
        specs=[[{"type": "bar"}], [{"type": "bar"}], [{"type": "bar"}]]
    )
    
    # Chart 1: Overall enrollment by year (using UNIQUE_ID for counting)
    yearly_counts = df.groupby('YEAR')['UNIQUE_ID'].nunique().sort_index()
    
    fig.add_trace(
        go.Bar(
            x=yearly_counts.index,
            y=yearly_counts.values,
            text=yearly_counts.values,
            textposition='auto',
            marker_color='skyblue',
            name='Total Students',
            hovertemplate='<b>Year:</b> %{x}<br>' +
                         '<b>Number of Students:</b> %{y}<br>' +
                         '<extra></extra>'
        ),
        row=1, col=1
    )
    
    # Assign a color to each course (for both charts)
    all_courses = sorted(df['COURSE'].dropna().unique())
    colors = px.colors.qualitative.Set3 * ((len(all_courses) // len(px.colors.qualitative.Set3)) + 1)
    course_color_map = {course: colors[i] for i, course in enumerate(all_courses)}

    # Chart 2: Course-wise enrollment by year (using UNIQUE_ID for counting)
    course_yearly_counts = df.groupby(['YEAR', 'COURSE'])['UNIQUE_ID'].nunique().reset_index(name='count')
    # Define the correct chronological order for years
    year_order = sorted(df['YEAR'].dropna().unique(), key=lambda x: (int(x.split('-')[0]), int(x.split('-')[1])))
    course_yearly_counts['YEAR'] = pd.Categorical(course_yearly_counts['YEAR'], categories=year_order, ordered=True)
    course_yearly_counts = course_yearly_counts.sort_values(['COURSE', 'YEAR'])

    for i, course in enumerate(all_courses):
        course_data = course_yearly_counts[course_yearly_counts['COURSE'] == course]
        # Ensure x is in the correct order and all years are present
        course_data = course_data.set_index('YEAR').reindex(year_order).reset_index()
        showlegend = True  # Only show legend for the first trace of each course
        fig.add_trace(
            go.Bar(
                x=course_data['YEAR'],
                y=course_data['count'].fillna(0),
                name=course,
                text=course_data['count'].fillna(0),
                textposition='auto',
                marker_color=course_color_map[course],
                legendgroup=course,
                showlegend=showlegend,
                hovertemplate='<b>Year:</b> %{x}<br>' +
                             f'<b>Course:</b> {course}<br>' +
                             '<b>Number of Students:</b> %{y}<br>' +
                             '<extra></extra>'
            ),
            row=2, col=1
        )

    # Chart 3: Course popularity by duration (using UNIQUE_ID for counting)
    course_duration_counts = df.groupby(['COURSE', 'DURATION'])['UNIQUE_ID'].nunique().reset_index(name='count')
    course_duration_counts['COURSE_DURATION'] = course_duration_counts['COURSE'] + ' (' + course_duration_counts['DURATION'] + ')'
    course_duration_counts = course_duration_counts.sort_values('count', ascending=False)

    # Add a bar for each course+duration, using the course color, legend shows course+duration
    for _, row in course_duration_counts.iterrows():
        fig.add_trace(
            go.Bar(
                x=[row['COURSE_DURATION']],
                y=[row['count']],
                text=[row['count']],
                textposition='auto',
                marker_color=course_color_map[row['COURSE']],
                name=row['COURSE_DURATION'],
                legendgroup=row['COURSE'],
                showlegend=True,
                hovertemplate='<b>Course:</b> %{x}<br>' +
                             '<b>Number of Students:</b> %{y}<br>' +
                             '<extra></extra>'
            ),
            row=3, col=1
        )
    
    # Update layout
    fig.update_layout(
        title='Student Enrollment Analysis Dashboard',
        template='plotly_white',
        height=1200,
        showlegend=True,
        barmode='group'
    )
    
    # Update x-axis and y-axis labels
    fig.update_xaxes(title_text="Academic Year", row=1, col=1)
    fig.update_yaxes(title_text="Number of Students", row=1, col=1)
    fig.update_xaxes(title_text="Academic Year", row=2, col=1)
    fig.update_yaxes(title_text="Number of Students", row=2, col=1)
    fig.update_xaxes(title_text="Course (Duration)", row=3, col=1)
    fig.update_yaxes(title_text="Number of Students", row=3, col=1)
    
    # Rotate x-axis labels for the third chart to prevent overlap
    fig.update_xaxes(tickangle=45, row=3, col=1)
    
    # In the layout, set category_orders for the x-axis of row 2
    fig.update_xaxes(categoryorder='array', categoryarray=year_order, row=2, col=1)
    
    # Save the combined chart
    fig.write_html('student_enrollment_dashboard.html')
    print("Combined dashboard saved as 'student_enrollment_dashboard.html'")

def create_student_enrollment_charts():
    """
    Create interactive bar charts for student enrollment analysis
    """
    # Combine all dataframes
    all_data = []
    for df in sheet_data.values():
        all_data.append(df)
    
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # Remove rows where YEAR is None
    combined_df = combined_df.dropna(subset=['YEAR'])
    
    # Check combination uniqueness
    is_unique = check_combination_uniqueness(combined_df)
    
    # Create combined chart
    create_combined_charts(combined_df)

if __name__ == "__main__":
    create_student_enrollment_charts() 