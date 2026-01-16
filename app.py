import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Employee Classification Matrix", page_icon="📊")

st.title("📊 Employee Classification System")
st.write("Upload an Excel file with columns: **Name, Discipline, Attendance**")

uploaded = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded:

    df = pd.read_excel(uploaded)

    # Validate required columns
    required_cols = {"Name", "Discipline", "Attendance"}
    if not required_cols.issubset(df.columns):
        st.error(f"Excel must contain columns: {required_cols}")
        st.stop()

    # Create matrix categories
    matrix = {
        (0, 0): [],  # Att=0, Disc=1 (Top-left)
        (0, 1): [],  # Att=1, Disc=1 (Top-right)
        (1, 0): [],  # Att=0, Disc=0 (Bottom-left)
        (1, 1): []   # Att=1, Disc=0 (Bottom-right)
    }

    for _, row in df.iterrows():
        att = row['Attendance']
        disc = row['Discipline']
        name = row['Name']

        if att == 0 and disc == 1:
            matrix[(0, 0)].append(name)
        elif att == 1 and disc == 1:
            matrix[(0, 1)].append(name)
        elif att == 0 and disc == 0:
            matrix[(1, 0)].append(name)
        else:
            matrix[(1, 1)].append(name)

    # Matrix display DataFrame
    matrix_df = pd.DataFrame([
        [', '.join(matrix[(0, 0)]), ', '.join(matrix[(0, 1)])],
        [', '.join(matrix[(1, 0)]), ', '.join(matrix[(1, 1)])]
    ],
    index=[1, 0],
    columns=[0, 1])

    matrix_df.index.name = "Discipline"
    matrix_df.columns.name = "Attendance"

    # Count matrix
    count_df = pd.DataFrame([
        [len(matrix[(0, 0)]), len(matrix[(0, 1)])],
        [len(matrix[(1, 0)]), len(matrix[(1, 1)])]
    ],
    index=[1, 0],
    columns=[0, 1])

    count_df.index.name = "Discipline"
    count_df.columns.name = "Attendance"

    # Classification
    def classify(row):
        if row['Attendance'] == 0 and row['Discipline'] == 1:
            return 'Low Attendance, Good Discipline'
        elif row['Attendance'] == 1 and row['Discipline'] == 1:
            return 'Good Attendance, Good Discipline'
        elif row['Attendance'] == 0 and row['Discipline'] == 0:
            return 'Low Attendance, Poor Discipline'
        else:
            return 'Good Attendance, Poor Discipline'

    df['Classification'] = df.apply(classify, axis=1)

    # Summary
    summary = pd.DataFrame({
        'Category': [
            'Total Students',
            'Good Attendance (1)',
            'Low Attendance (0)',
            'Good Discipline (1)',
            'Poor Discipline (0)',
            'Both Good (Att=1, Disc=1)',
            'Both Poor (Att=0, Disc=0)'
        ],
        'Count': [
            len(df),
            len(df[df['Attendance'] == 1]),
            len(df[df['Attendance'] == 0]),
            len(df[df['Discipline'] == 1]),
            len(df[df['Discipline'] == 0]),
            len(df[(df['Attendance'] == 1) & (df['Discipline'] == 1)]),
            len(df[(df['Attendance'] == 0) & (df['Discipline'] == 0)])
        ]
    })

    st.success("Processing complete!")

    st.subheader("Classification Matrix")
    st.dataframe(matrix_df)

    st.subheader("Count Matrix")
    st.dataframe(count_df)

    st.subheader("Student Details with Classification")
    st.dataframe(df)

    # Generate excel with 4 sheets
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Student Details")
        matrix_df.to_excel(writer, sheet_name="Classification Matrix")
        count_df.to_excel(writer, sheet_name="Count Matrix")
        summary.to_excel(writer, index=False, sheet_name="Summary")

    st.download_button(
        label="📁 Download Result Excel",
        data=output.getvalue(),
        file_name="student_classification_output.xlsx"
    )
