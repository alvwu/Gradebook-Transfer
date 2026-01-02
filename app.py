import streamlit as st
import pandas as pd
import zipfile
import io
import tempfile
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows


def parse_numbers_file(uploaded_file):
    """Parse a .numbers file and return a pandas DataFrame."""
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save uploaded file to temp location
        temp_numbers_path = os.path.join(temp_dir, "temp.numbers")
        with open(temp_numbers_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # .numbers files are zip archives
        with zipfile.ZipFile(temp_numbers_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Find and parse the tables
        tables_dir = os.path.join(temp_dir, "Index", "Tables")

        if not os.path.exists(tables_dir):
            # Try alternative structure
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if file.endswith('.iwa'):
                        pass  # .iwa files need special parsing
            raise ValueError("Could not find tables in the Numbers file. Please try exporting as CSV.")

        # For .numbers files, we need to use the numbers-parser library
        try:
            from numbers_parser import Document
            doc = Document(temp_numbers_path)
            sheets = doc.sheets
            if not sheets:
                raise ValueError("No sheets found in the Numbers file.")

            # Get the first table from the first sheet
            table = sheets[0].tables[0]

            # Extract data
            data = []
            headers = []

            for row_num, row in enumerate(table.iter_rows()):
                row_data = []
                for cell in row:
                    row_data.append(cell.value if cell.value is not None else "")

                if row_num == 0:
                    headers = row_data
                else:
                    data.append(row_data)

            df = pd.DataFrame(data, columns=headers)
            return df

        except ImportError:
            raise ImportError("The 'numbers-parser' library is required. Please install it with: pip install numbers-parser")


def categorize_columns(columns, category_keywords):
    """Categorize columns based on keywords."""
    categorized = {}
    uncategorized = []

    for col in columns:
        col_lower = col.lower()
        found_category = None

        for category, keywords in category_keywords.items():
            for keyword in keywords:
                if keyword.lower() in col_lower:
                    found_category = category
                    break
            if found_category:
                break

        if found_category:
            if found_category not in categorized:
                categorized[found_category] = []
            categorized[found_category].append(col)
        else:
            uncategorized.append(col)

    return categorized, uncategorized


def create_student_excel(df, name_column, category_keywords, show_category_averages):
    """Create an Excel file with each student on their own sheet."""
    output = io.BytesIO()
    wb = Workbook()

    # Remove default sheet
    wb.remove(wb.active)

    # Get grade columns (all columns except the name column)
    grade_columns = [col for col in df.columns if col != name_column]

    # Categorize columns
    categorized, uncategorized = categorize_columns(grade_columns, category_keywords)

    # Style definitions
    header_font = Font(bold=True, size=12)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(bold=True, size=12, color="FFFFFF")
    category_fill = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
    category_font = Font(bold=True, size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center')

    # Create a sheet for each student
    for _, row in df.iterrows():
        student_name = str(row[name_column]).strip()

        # Sanitize sheet name (Excel has restrictions)
        safe_name = student_name[:31]  # Max 31 chars
        safe_name = ''.join(c for c in safe_name if c not in '[]:*?/\\')
        if not safe_name:
            safe_name = f"Student_{_}"

        # Handle duplicate sheet names
        original_name = safe_name
        counter = 1
        while safe_name in wb.sheetnames:
            safe_name = f"{original_name[:28]}_{counter}"
            counter += 1

        ws = wb.create_sheet(title=safe_name)

        # Add student name header
        ws['A1'] = "Student Name:"
        ws['B1'] = student_name
        ws['A1'].font = header_font
        ws['B1'].font = Font(bold=True, size=14)

        current_row = 3
        all_grades = []
        category_averages = {}

        # Add headers
        ws.cell(row=current_row, column=1, value="Assignment")
        ws.cell(row=current_row, column=2, value="Score")
        ws.cell(row=current_row, column=1).font = header_font_white
        ws.cell(row=current_row, column=2).font = header_font_white
        ws.cell(row=current_row, column=1).fill = header_fill
        ws.cell(row=current_row, column=2).fill = header_fill
        ws.cell(row=current_row, column=1).border = border
        ws.cell(row=current_row, column=2).border = border
        ws.cell(row=current_row, column=1).alignment = center_align
        ws.cell(row=current_row, column=2).alignment = center_align

        current_row += 1

        # Add categorized grades
        for category, columns in categorized.items():
            # Category header
            ws.cell(row=current_row, column=1, value=category.upper())
            ws.cell(row=current_row, column=1).font = category_font
            ws.cell(row=current_row, column=1).fill = category_fill
            ws.cell(row=current_row, column=2).fill = category_fill
            ws.cell(row=current_row, column=1).border = border
            ws.cell(row=current_row, column=2).border = border
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            current_row += 1

            category_grades = []
            for col in columns:
                try:
                    grade = float(row[col]) if pd.notna(row[col]) and row[col] != '' else 0
                except (ValueError, TypeError):
                    grade = 0

                ws.cell(row=current_row, column=1, value=col)
                ws.cell(row=current_row, column=2, value=grade)
                ws.cell(row=current_row, column=1).border = border
                ws.cell(row=current_row, column=2).border = border
                ws.cell(row=current_row, column=2).alignment = center_align

                category_grades.append(grade)
                all_grades.append(grade)
                current_row += 1

            # Calculate category average
            if category_grades:
                category_averages[category] = sum(category_grades) / len(category_grades)

        # Add uncategorized grades
        if uncategorized:
            ws.cell(row=current_row, column=1, value="OTHER")
            ws.cell(row=current_row, column=1).font = category_font
            ws.cell(row=current_row, column=1).fill = category_fill
            ws.cell(row=current_row, column=2).fill = category_fill
            ws.cell(row=current_row, column=1).border = border
            ws.cell(row=current_row, column=2).border = border
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            current_row += 1

            other_grades = []
            for col in uncategorized:
                try:
                    grade = float(row[col]) if pd.notna(row[col]) and row[col] != '' else 0
                except (ValueError, TypeError):
                    grade = 0

                ws.cell(row=current_row, column=1, value=col)
                ws.cell(row=current_row, column=2, value=grade)
                ws.cell(row=current_row, column=1).border = border
                ws.cell(row=current_row, column=2).border = border
                ws.cell(row=current_row, column=2).alignment = center_align

                other_grades.append(grade)
                all_grades.append(grade)
                current_row += 1

            if other_grades:
                category_averages["Other"] = sum(other_grades) / len(other_grades)

        current_row += 1

        # Add category averages if enabled
        if show_category_averages and category_averages:
            ws.cell(row=current_row, column=1, value="CATEGORY AVERAGES")
            ws.cell(row=current_row, column=1).font = header_font_white
            ws.cell(row=current_row, column=1).fill = header_fill
            ws.cell(row=current_row, column=2).fill = header_fill
            ws.cell(row=current_row, column=1).border = border
            ws.cell(row=current_row, column=2).border = border
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            current_row += 1

            for category, avg in category_averages.items():
                ws.cell(row=current_row, column=1, value=category)
                ws.cell(row=current_row, column=2, value=round(avg, 4))
                ws.cell(row=current_row, column=1).border = border
                ws.cell(row=current_row, column=2).border = border
                ws.cell(row=current_row, column=2).alignment = center_align
                current_row += 1

            current_row += 1

        # Add overall final grade
        if all_grades:
            final_grade = sum(all_grades) / len(all_grades)
            ws.cell(row=current_row, column=1, value="FINAL GRADE")
            ws.cell(row=current_row, column=2, value=round(final_grade, 4))
            ws.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            ws.cell(row=current_row, column=2).font = Font(bold=True, size=12)
            ws.cell(row=current_row, column=1).fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
            ws.cell(row=current_row, column=2).fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
            ws.cell(row=current_row, column=1).border = border
            ws.cell(row=current_row, column=2).border = border
            ws.cell(row=current_row, column=2).alignment = center_align

        # Adjust column widths
        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 15

    wb.save(output)
    output.seek(0)
    return output


def main():
    st.set_page_config(
        page_title="GradeBook Transfer",
        page_icon="📚",
        layout="wide"
    )

    # Custom CSS for better styling
    st.markdown("""
        <style>
        .main-header {
            font-size: 2.5rem;
            font-weight: bold;
            color: #1E3A5F;
            text-align: center;
            margin-bottom: 0.5rem;
        }
        .sub-header {
            font-size: 1.1rem;
            color: #666;
            text-align: center;
            margin-bottom: 2rem;
        }
        .success-box {
            padding: 1rem;
            border-radius: 0.5rem;
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            margin: 1rem 0;
        }
        .info-box {
            padding: 1rem;
            border-radius: 0.5rem;
            background-color: #e7f3ff;
            border: 1px solid #b6d4fe;
            margin: 1rem 0;
        }
        .stButton>button {
            width: 100%;
        }
        </style>
    """, unsafe_allow_html=True)

    # Header
    st.markdown('<p class="main-header">📚 GradeBook Transfer</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Convert your Numbers gradebook into organized Excel files with individual student sheets</p>', unsafe_allow_html=True)

    st.divider()

    # Sidebar for settings
    with st.sidebar:
        st.header("⚙️ Settings")

        st.subheader("Grade Display Options")
        show_category_averages = st.checkbox(
            "Show category averages",
            value=False,
            help="Display average scores for each category (Exams, Assignments, etc.) in addition to the overall final grade"
        )

        st.divider()

        st.subheader("Category Keywords")
        st.caption("Add keywords to categorize grade columns. Columns containing these keywords will be grouped together.")

        # Initialize session state for categories
        if 'categories' not in st.session_state:
            st.session_state.categories = {
                "Exams": ["exam", "test", "midterm", "final"],
                "Assignments": ["assignment", "homework", "hw"],
                "Participation": ["participation", "attendance"],
                "Quizzes": ["quiz"]
            }

        # Display existing categories
        categories_to_remove = []
        for category in list(st.session_state.categories.keys()):
            with st.expander(f"📁 {category}", expanded=False):
                keywords = st.session_state.categories[category]
                new_keywords = st.text_area(
                    "Keywords (one per line)",
                    value="\n".join(keywords),
                    key=f"keywords_{category}",
                    height=100
                )
                st.session_state.categories[category] = [k.strip() for k in new_keywords.split("\n") if k.strip()]

                if st.button(f"🗑️ Remove {category}", key=f"remove_{category}"):
                    categories_to_remove.append(category)

        # Remove marked categories
        for cat in categories_to_remove:
            del st.session_state.categories[cat]
            st.rerun()

        st.divider()

        # Add new category
        st.subheader("Add New Category")
        new_category_name = st.text_input("Category name", placeholder="e.g., Projects")
        new_category_keywords = st.text_input("Keywords (comma-separated)", placeholder="e.g., project, proj")

        if st.button("➕ Add Category"):
            if new_category_name and new_category_keywords:
                keywords_list = [k.strip() for k in new_category_keywords.split(",") if k.strip()]
                if keywords_list:
                    st.session_state.categories[new_category_name] = keywords_list
                    st.success(f"Added category: {new_category_name}")
                    st.rerun()
            else:
                st.warning("Please enter both category name and keywords")

    # Main content area
    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📤 Upload Your Numbers File")

        uploaded_file = st.file_uploader(
            "Drag and drop your .numbers file here",
            type=["numbers"],
            help="Upload your Numbers gradebook file. The first row should contain headers (student name and grade categories)."
        )

    with col2:
        st.subheader("📋 How It Works")
        st.markdown("""
        1. **Upload** your .numbers gradebook
        2. **Select** the column containing student names
        3. **Review** the detected categories
        4. **Download** the organized Excel file
        """)

    if uploaded_file is not None:
        st.divider()

        with st.spinner("🔄 Parsing Numbers file..."):
            try:
                df = parse_numbers_file(uploaded_file)
                st.success("✅ File parsed successfully!")

                # Show preview
                st.subheader("📊 Data Preview")
                st.dataframe(df.head(10), use_container_width=True)

                st.info(f"📈 Found **{len(df)}** students and **{len(df.columns)}** columns")

                # Select name column
                st.subheader("🏷️ Select Name Column")
                name_column = st.selectbox(
                    "Which column contains student names?",
                    options=df.columns.tolist(),
                    index=0
                )

                # Show category detection preview
                st.subheader("🗂️ Category Detection Preview")
                grade_columns = [col for col in df.columns if col != name_column]
                categorized, uncategorized = categorize_columns(grade_columns, st.session_state.categories)

                col1, col2 = st.columns(2)

                with col1:
                    st.markdown("**Categorized Columns:**")
                    for category, columns in categorized.items():
                        with st.expander(f"{category} ({len(columns)} items)"):
                            for col in columns:
                                st.write(f"• {col}")

                with col2:
                    st.markdown("**Uncategorized Columns:**")
                    if uncategorized:
                        for col in uncategorized:
                            st.write(f"• {col}")
                        st.caption("💡 Add keywords in the sidebar to categorize these columns")
                    else:
                        st.write("All columns are categorized!")

                st.divider()

                # Generate Excel
                st.subheader("📥 Generate Excel File")

                if st.button("🚀 Generate Excel File", type="primary", use_container_width=True):
                    with st.spinner("🔄 Creating Excel file with individual student sheets..."):
                        excel_output = create_student_excel(
                            df,
                            name_column,
                            st.session_state.categories,
                            show_category_averages
                        )

                        st.success("✅ Excel file generated successfully!")

                        # Summary
                        st.markdown(f"""
                        <div class="success-box">
                            <h4>✨ Generation Complete!</h4>
                            <p>Created <strong>{len(df)}</strong> individual student sheets</p>
                            <p>Each sheet contains:</p>
                            <ul>
                                <li>Student name</li>
                                <li>All assignments organized by category</li>
                                <li>Scores (0 or 1)</li>
                                {"<li>Category averages</li>" if show_category_averages else ""}
                                <li>Overall final grade (average of all scores)</li>
                            </ul>
                        </div>
                        """, unsafe_allow_html=True)

                        # Download button
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_output,
                            file_name="gradebook_transfer.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            except ImportError as e:
                st.error(f"❌ {str(e)}")
                st.info("💡 Run this command in your terminal: `pip install numbers-parser`")
            except Exception as e:
                st.error(f"❌ Error parsing file: {str(e)}")
                st.info("💡 Make sure your Numbers file has a table with headers in the first row")

    # Footer
    st.divider()
    st.markdown(
        "<p style='text-align: center; color: #888;'>GradeBook Transfer | Made for educators</p>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
