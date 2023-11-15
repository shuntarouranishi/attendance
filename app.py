import streamlit as st
import pandas as pd
import pulp
import holidays
from datetime import datetime
from io import BytesIO

# Define jp_holidays in the global scope
jp_holidays = holidays.Japan()

# Function to check if a date is a weekend or a holiday
def is_holiday_or_weekend(date):
    return date.weekday() in [5, 6] or date in jp_holidays

# Streamlit app title and file uploader
st.title('シフト作成アプリ')
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Reading the uploaded Excel file
    data = pd.read_excel(uploaded_file)
    days_in_month = 30  # Number of days in the target month

    # Initialize lists for rest1, rest2, and morning_shift
    rest1, rest2, morning_shift = [], [], []

    # Process the data
    for _, row in data.iterrows():
        rest1_days = [0] * days_in_month
        rest2_days = [0] * days_in_month
        morning_shift_days = [0] * days_in_month

        # Process for 特別休 and 有給 (rest1)
        for col in ["特別休", "有給1", "有給2", "有給3", "有給4", "有給5", "有給6", "有給7", "有給8", "有給9"]:
            if col in data.columns and not pd.isna(row[col]):
                rest1_days[row[col].day - 1] = 1

        # Process for 希望休 (rest2) - considering weekends and holidays
        for col in ["希望休1", "希望休2", "希望休3"]:
            if col in data.columns and not pd.isna(row[col]):
                date = datetime(2023, 11, row[col].day)
                if is_holiday_or_weekend(date):
                    rest2_days[row[col].day - 1] = -5

        # Process for 朝番 (morning shift availability)
        for day in range(1, days_in_month + 1):
            if row.get(f'朝番{day}', 0) == 1:
                morning_shift_days[day - 1] = 1

        rest1.append(rest1_days)
        rest2.append(rest2_days)
        morning_shift.append(morning_shift_days)

    # Optimization problem
    prob = pulp.LpProblem('Maximize_Score', pulp.LpMaximize)

    # Define variables
    x = pulp.LpVariable.dicts('x', (range(len(data)), range(days_in_month)), cat='Binary')

    # Objective function
    score = []
    for i in range(len(data)):
        for j in range(days_in_month):
            point = 10 if is_holiday_or_weekend(datetime(2023, 11, j + 1)) else 7
            point += rest2[i][j]
            score.append(point * x[i][j])
    prob += pulp.lpSum(score)

    # Constraints
    for i in range(len(data)):
        prob += pulp.lpSum(x[i][j] for j in range(days_in_month)) == 20

    for j in range(days_in_month):
        prob += pulp.lpSum(x[i][j] for i in range(len(data))) >= 2

    for i in range(len(data)):
        for j in range(days_in_month):
            if rest1[i][j] == 1:
                prob += x[i][j] == 0

    # Constraint for morning shift
    for j in range(days_in_month):
        morning_shift_available = [i for i in range(len(data)) if morning_shift[i][j] == 1]
        if morning_shift_available:
            prob += pulp.lpSum(x[i][j] for i in morning_shift_available) >= 1

    # Solve the problem
    prob.solve()

    # Check if a solution is found
    if pulp.LpStatus[prob.status] == "Optimal":
        # Process the results
        weekday_holiday_header = []
        for day in range(1, days_in_month + 1):
            date = datetime(2023, 11, day)
            if date in jp_holidays:
                holiday_name = jp_holidays[date]
                weekday_holiday_header.append(f'{date.day} ({holiday_name})')
            else:
                weekday = date.strftime('%a')
                weekday_holiday_header.append(f'{date.day} ({weekday})')

        # DataFrame for shift schedule
        shift_df_with_header = pd.DataFrame([weekday_holiday_header], columns=[f'{j + 1}' for j in range(days_in_month)])
        output_data = {f'{j + 1}': [] for j in range(days_in_month)}
        for i in range(len(data)):
            for j in range(days_in_month):
                if morning_shift[i][j] == 1 and x[i][j].varValue == 1:
                    output_data[f'{j + 1}'].append(2)
                else:
                    cell_value = 0 if x[i][j].varValue is None else int(x[i][j].varValue)
                    output_data[f'{j + 1}'].append(cell_value)

        shift_df_with_dates = pd.DataFrame(output_data)
        final_shift_df = pd.concat([shift_df_with_header, shift_df_with_dates], ignore_index=True)
        final_shift_df.insert(0, '名前', [''] + data['申請者'].tolist())

        # Convert final DataFrame to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_shift_df.to_excel(writer, index=False)
        output.seek(0)

        # Download link for the processed Excel file
        st.download_button(label="Download Processed Excel File",
                           data=output,
                           file_name="processed_output.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:
        st.error("No optimal solution found. Please check the input data.")

# Additional features like error handling, user instructions, etc., can be added here
