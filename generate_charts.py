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
    # Create subplots with 7 rows and 1 column
    fig = make_subplots(
        rows=7, cols=1,
        subplot_titles=[
            'Student Enrollment by Year',
            'Student Enrollment by Year and Course',
            'Course Popularity by Duration',
            'Gender Distribution',
            'Gender Distribution per Year',
            'Total vs Employed Students per Year',
            'Total vs Employed Students per Year and Course'
        ],
        vertical_spacing=0.08,  # Adjusted for 7 charts
        specs=[[{"type": "bar"}], [{"type": "bar"}], [{"type": "bar"}], [{"type": "pie"}], [{"type": "bar"}], [{"type": "bar"}], [{"type": "bar"}]]
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
    
    # Define gender color map
    gender_color_map = {
        'Male': '#4F6BED',    # blue
        'Female': '#F653A6'   # pink
    }

    # Chart 4: Gender distribution (using UNIQUE_ID for counting)
    gender_counts = df.groupby('GENDER')['UNIQUE_ID'].nunique().reset_index(name='count')
    gender_counts = gender_counts[gender_counts['GENDER'] != 'Unknown']  # Optionally exclude 'Unknown'

    # Only show legend for gender in pie chart, not in bar chart (handled below)
    fig.add_trace(
        go.Pie(
            labels=gender_counts['GENDER'],
            values=gender_counts['count'],
            textinfo='label+percent',
            hoverinfo='label+value+percent',
            name='Gender Distribution',
            hole=0.4,
            marker=dict(
                colors=[gender_color_map.get(g, '#CCCCCC') for g in gender_counts['GENDER']],
                line=dict(color='#000000', width=1)
            ),
            sort=False,
            legendgroup='gender',
            showlegend=True
        ),
        row=4, col=1
    )
    
    # Chart 5: Gender distribution per year (bar chart, using UNIQUE_ID for counting)
    gender_year_counts = df.groupby(['YEAR', 'GENDER'])['UNIQUE_ID'].nunique().reset_index(name='count')
    gender_year_counts = gender_year_counts[gender_year_counts['GENDER'] != 'Unknown']  # Optionally exclude 'Unknown'
    gender_list = gender_year_counts['GENDER'].unique()

    for i, gender in enumerate(gender_list):
        gender_data = gender_year_counts[gender_year_counts['GENDER'] == gender]
        fig.add_trace(
            go.Bar(
                x=gender_data['YEAR'],
                y=gender_data['count'],
                name=gender,
                text=gender_data['count'],
                textposition='auto',
                marker_color=gender_color_map.get(gender, '#CCCCCC'),
                legendgroup='gender',
                showlegend=False,  # Only show legend in pie chart
                hovertemplate='<b>Year:</b> %{x}<br>' +
                             f'<b>Gender:</b> {gender}<br>' +
                             '<b>Number of Students:</b> %{y}<br>' +
                             '<extra></extra>'
            ),
            row=5, col=1
        )

    # Chart 6: Total vs Employed Students per Year
    total_per_year = df.groupby('YEAR')['UNIQUE_ID'].nunique().reset_index(name='Total Students')
    employed_per_year = df[df['EMPLOYMENT_STATUS'] == 'Employed'].groupby('YEAR')['UNIQUE_ID'].nunique().reset_index(name='Employed Students')
    merged_year = pd.merge(total_per_year, employed_per_year, on='YEAR', how='left').fillna(0)
    x_years = merged_year['YEAR']
    fig.add_trace(
        go.Bar(
            x=x_years,
            y=merged_year['Total Students'],
            name='Total Students',
            marker_color='#7fc7e3',
            text=merged_year['Total Students'],
            textposition='auto',
            legendgroup='employment',
            showlegend=True,
            hovertemplate='<b>Year:</b> %{x}<br>Total Students: %{y}<extra></extra>'
        ),
        row=6, col=1
    )
    fig.add_trace(
        go.Bar(
            x=x_years,
            y=merged_year['Employed Students'],
            name='Employed Students',
            marker_color='#2ca02c',
            text=merged_year['Employed Students'],
            textposition='auto',
            legendgroup='employment',
            showlegend=True,
            hovertemplate='<b>Year:</b> %{x}<br>Employed Students: %{y}<extra></extra>'
        ),
        row=6, col=1
    )

    # Chart 7: Employed Students per Year and Course (no total, sorted by year)
    employed_per_year_course = df[df['EMPLOYMENT_STATUS'] == 'Employed'].groupby(['YEAR', 'COURSE'])['UNIQUE_ID'].nunique().reset_index(name='Employed Students')
    # Sort years as in chart 2 and fill missing years with 0 for each course
    employed_per_year_course['YEAR'] = pd.Categorical(employed_per_year_course['YEAR'], categories=year_order, ordered=True)
    employed_per_year_course = employed_per_year_course.sort_values(['COURSE', 'YEAR'])
    for course in all_courses:
        course_data = employed_per_year_course[employed_per_year_course['COURSE'] == course].set_index('YEAR').reindex(year_order).reset_index()
        fig.add_trace(
            go.Bar(
                x=course_data['YEAR'],
                y=course_data['Employed Students'].fillna(0),
                name=f'Employed Students - {course}',
                marker_color=course_color_map[course],
                text=course_data['Employed Students'].fillna(0),
                textposition='auto',
                legendgroup=f'employment_{course}',
                showlegend=False,
                hovertemplate=f'<b>Year:</b> %{{x}}<br>Employed Students ({course}): %{{y}}<extra></extra>'
            ),
            row=7, col=1
        )

    # Update layout for 7 rows
    fig.update_layout(
        title='Student Enrollment Analysis Dashboard',
        template='plotly_white',
        height=2500,  # Increased height for 7 charts
        showlegend=False,
        barmode='group',
        margin=dict(l=60, r=260, t=80, b=80)  # Increased right margin
    )

    # Update x-axis and y-axis labels for all charts
    fig.update_xaxes(title_text="Academic Year", row=1, col=1)
    fig.update_yaxes(title_text="Number of Students", row=1, col=1)
    fig.update_xaxes(title_text="Academic Year", row=2, col=1)
    fig.update_yaxes(title_text="Number of Students", row=2, col=1)
    fig.update_xaxes(title_text="Course (Duration)", row=3, col=1)
    fig.update_yaxes(title_text="Number of Students", row=3, col=1)
    fig.update_xaxes(showticklabels=False, row=4, col=1)
    fig.update_yaxes(showticklabels=False, row=4, col=1)
    fig.update_xaxes(title_text="Academic Year", row=5, col=1)
    fig.update_yaxes(title_text="Number of Students", row=5, col=1)
    fig.update_xaxes(title_text="Academic Year", row=6, col=1)
    fig.update_yaxes(title_text="Number of Students", row=6, col=1)
    fig.update_xaxes(title_text="Academic Year", row=7, col=1)
    fig.update_yaxes(title_text="Number of Students", row=7, col=1)
    
    # Rotate x-axis labels for the third and fifth chart to prevent overlap
    fig.update_xaxes(tickangle=45, row=3, col=1)
    fig.update_xaxes(tickangle=45, row=5, col=1)
    
    # In the layout, set category_orders for the x-axis of row 2 and 5
    fig.update_xaxes(categoryorder='array', categoryarray=year_order, row=2, col=1)
    fig.update_xaxes(categoryorder='array', categoryarray=year_order, row=5, col=1)
    
    # Hide the global legend
    fig.update_layout(showlegend=False)

    # Add custom legend for Chart 2 (Course-wise enrollment by year)
    course_legend_text = "<b>Course</b><br>" + "<br>".join(
        f"<span style='color:{course_color_map[c]}'>&#9632;</span> {c}" for c in all_courses
    )
    fig.add_annotation(
        dict(
            x=1.01,
            y=0.85,
            xref='paper',
            yref='paper',
            text=course_legend_text,
            showarrow=False,
            align='left',
            xanchor='left',
            yanchor='top',
            font=dict(size=13),
            bordercolor="#cccccc",
            borderwidth=1,
            bgcolor="#fff"
        )
    )

    # Add custom legend for Chart 3 (Course Popularity by Duration)
    course_duration_legend_text = "<b>Course</b><br>" + "<br>".join(
        f"<span style='color:{course_color_map[c]}'>&#9632;</span> {c}" for c in all_courses
    )
    fig.add_annotation(
        dict(
            x=1.01,
            y=0.68,
            xref='paper',
            yref='paper',
            text=course_duration_legend_text,
            showarrow=False,
            align='left',
            xanchor='left',
            yanchor='top',
            font=dict(size=13),
            bordercolor="#cccccc",
            borderwidth=1,
            bgcolor="#fff"
        )
    )

    # Add custom legend for Chart 4 (Gender Pie)
    gender_legend_text = "<b>Gender</b><br>" + "<br>".join(
        f"<span style='color:{gender_color_map[g]}'>&#9632;</span> {g}" for g in gender_color_map.keys()
    )
    fig.add_annotation(
        dict(
            x=1.01,
            y=0.52,
            xref='paper',
            yref='paper',
            text=gender_legend_text,
            showarrow=False,
            align='left',
            xanchor='left',
            yanchor='top',
            font=dict(size=13),
            bordercolor="#cccccc",
            borderwidth=1,
            bgcolor="#fff"
        )
    )
    # Add custom legend for Chart 5 (Gender Bar)
    fig.add_annotation(
        dict(
            x=1.01,
            y=0.36,
            xref='paper',
            yref='paper',
            text=gender_legend_text,
            showarrow=False,
            align='left',
            xanchor='left',
            yanchor='top',
            font=dict(size=13),
            bordercolor="#cccccc",
            borderwidth=1,
            bgcolor="#fff"
        )
    )
    # Add custom legend for Chart 6 (Employment)
    employment_legend_text = "<b>Employment</b><br>" + \
        f"<span style='color:#7fc7e3'>&#9632;</span> Total Students<br>" + \
        f"<span style='color:#2ca02c'>&#9632;</span> Employed Students"
    fig.add_annotation(
        dict(
            x=1.01,
            y=0.22,
            xref='paper',
            yref='paper',
            text=employment_legend_text,
            showarrow=False,
            align='left',
            xanchor='left',
            yanchor='top',
            font=dict(size=13),
            bordercolor="#cccccc",
            borderwidth=1,
            bgcolor="#fff"
        )
    )
    # Add custom legend for Chart 7 (Employment by Course)
    employment_course_legend_text = "<b>Course</b><br>" + "<br>".join(
        f"<span style='color:{course_color_map[c]}'>&#9632;</span> {c}" for c in all_courses
    )
    fig.add_annotation(
        dict(
            x=1.01,
            y=0.04,
            xref='paper',
            yref='paper',
            text=employment_course_legend_text,
            showarrow=False,
            align='left',
            xanchor='left',
            yanchor='top',
            font=dict(size=13),
            bordercolor="#cccccc",
            borderwidth=1,
            bgcolor="#fff"
        )
    )

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