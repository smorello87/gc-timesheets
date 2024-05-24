import pandas as pd
from datetime import datetime, timedelta
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfObject
import numpy as np
import holidays
import random

# Load Excel file
excel_file = '2024adjuncts.xlsx'
df = pd.read_excel(excel_file, sheet_name='Sheet1', header=None)

# Extract names
names = df.iloc[3, [4, 5, 10, 11, 12, 13, 14, 15, 16, 19, 20, 21, 22]].values

# Extract pay dates and hours worked
pay_dates_week1 = df.iloc[6:19, 2].astype(str).values
pay_dates_week2 = df.iloc[6:19, 3].astype(str).values
hours_worked = df.iloc[6:19, [4, 5, 10, 11, 12, 13, 14, 15, 16, 19, 20, 21, 22]].values
pay_dates = df.iloc[6:19, 0].astype(str).values

# US holidays
us_holidays = holidays.US(years=2024)

# Function to ensure date format consistency
def ensure_date_format(date):
    if isinstance(date, datetime):
        return date
    if isinstance(date, pd.Timestamp):
        return date.to_pydatetime()
    elif isinstance(date, str):
        try:
            return datetime.strptime(date, '%Y-%m-%d')
        except ValueError:
            return pd.to_datetime(date).to_pydatetime()
    return date

# Function to get workdays excluding holidays
def get_workdays(start_date):
    workdays = []
    for i in range(5):  # For a 5-day workweek
        current_date = start_date + timedelta(days=i)
        # Check if the current date is a workday (not a weekend or a holiday)
        if current_date.weekday() < 5 and current_date.strftime('%Y-%m-%d') not in us_holidays:
            workdays.append(current_date)
        else:
            print(f"Skipping {current_date.strftime('%Y-%m-%d')} - it's a weekend or a holiday")
    return workdays

# Function to split hours across workdays
def split_hours_across_workdays(workdays, total_hours):
    hours_per_day = [0] * len(workdays)
    remaining_hours = total_hours

    for i in range(len(workdays)):
        # Convert workdays[i] to a string in the same format as us_holidays
        if workdays[i].strftime('%Y-%m-%d') in us_holidays:  # Skip if the day is a holiday
            continue
        if remaining_hours > 8:
            hours_per_day[i] = 8
            remaining_hours -= 8
        else:
            hours_per_day[i] = remaining_hours
            break  # No more hours to assign

    return hours_per_day

# Function to create pay period data with random in/out times
def create_pay_period_data_with_times(start_date, total_hours):
    start_date = ensure_date_format(start_date)
    dates = [(start_date + timedelta(days=i)).strftime('%Y-%m-%d') for i in range(7)]
    workdays = get_workdays(start_date)
    hours_per_workday = split_hours_across_workdays(workdays, total_hours)

    hours_week = [0] * 7  # Initialize with 0 for all days (Sunday to Saturday)

    # Assign hours to the corresponding days in hours_week
    for i, date in enumerate(dates):
        if date in [workday.strftime('%Y-%m-%d') for workday in workdays]:
            hours_week[i] = hours_per_workday[workdays.index(datetime.strptime(date, '%Y-%m-%d'))]

    times = []
    for hours, date in zip(hours_week, dates):
        if hours == 0:
            times.append(('', '', '', ''))  # No in/out times for non-working days and holidays
        else:
            am_in = datetime.strptime('9:00', '%H:%M')
            am_out = am_in + timedelta(hours=hours)
            if hours >= 6:
                lunch_start = am_in + timedelta(hours=hours // 2)
                lunch_end = lunch_start + timedelta(hours=1)
                pm_out = am_out + timedelta(hours=1)
                times.append((am_in.strftime('%I:%M %p'), pm_out.strftime('%I:%M %p'), lunch_start.strftime('%I:%M %p'), lunch_end.strftime('%I:%M %p')))
            else:
                times.append((am_in.strftime('%I:%M %p'), am_out.strftime('%I:%M %p'), '', ''))
    return dates, hours_week, times

# Function to fill PDF
def fill_pdf(input_pdf_path, output_pdf_path, data):
    template_pdf = PdfReader(input_pdf_path)
    if not hasattr(template_pdf.Root, 'AcroForm'):
        template_pdf.Root.AcroForm = PdfDict()
    template_pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject('true')))
    
    for page in template_pdf.pages:
        annotations = page['/Annots']
        if annotations:
            for annotation in annotations:
                field = annotation['/T']
                if field:
                    field_name = field[1:-1]
                    if field_name in data.keys():
                        annotation.update(PdfDict(V=data[field_name]))
                        annotation.update(PdfDict(AP=None))  # Keep the fields editable

    PdfWriter().write(output_pdf_path, template_pdf)

# Prepare and fill data for each individual
for idx, name in enumerate(names):
    for row in range(7, 20):
        pay_date_week1 = ensure_date_format(pay_dates_week1[row - 7])
        pay_date_week2 = ensure_date_format(pay_dates_week2[row - 7])
        pay_date = ensure_date_format(pay_dates[row - 7]).strftime('%Y-%m-%d')
        total_hours = hours_worked[row - 7, idx]
        
        # Skip if no hours worked or if any value is NaN
        if total_hours == 0 or np.isnan(total_hours):
            continue
        
        # Split total hours evenly if needed
        half_hours = total_hours / 2
        if total_hours % 2 == 1:
            hours_week1 = half_hours + 0.5
            hours_week2 = half_hours - 0.5
        else:
            hours_week1 = half_hours
            hours_week2 = half_hours

        dates_week1, hours_week1, times_week1 = create_pay_period_data_with_times(pay_date_week1, hours_week1)
        dates_week2, hours_week2, times_week2 = create_pay_period_data_with_times(pay_date_week2, hours_week2)

        total_hours_week1 = sum(hours_week1)
        total_hours_week2 = sum(hours_week2)
        total_hours_period = total_hours_week1 + total_hours_week2

        data = {
            'Name': name,
            'Title': 'NTA V',
            'Employee Signature': name,
            'Supervisor': 'Steve Everett/MB',
            'Supervisor Signature': 'Steve Everett/MB',
            'Pay Date': pay_date,
            'DateSunday': dates_week1[0],
            'DateMonday': dates_week1[1],
            'DateTuesday': dates_week1[2],
            'DateWednesday': dates_week1[3],
            'DateThursday': dates_week1[4],
            'DateFriday': dates_week1[5],
            'DateSaturday': dates_week1[6],
            'Hours WorkedSunday': str(hours_week1[0]),
            'Hours WorkedMonday': str(hours_week1[1]),
            'Hours WorkedTuesday': str(hours_week1[2]),
            'Hours WorkedWednesday': str(hours_week1[3]),
            'Hours WorkedThursday': str(hours_week1[4]),
            'Hours WorkedFriday': str(hours_week1[5]),
            'Hours WorkedSaturday': str(hours_week1[6]),
            'AM InMonday': times_week1[1][0],
            'PM OutMonday': times_week1[1][1],
            'OutMonday': times_week1[1][2],
            'InMonday': times_week1[1][3],
            'AM InTuesday': times_week1[2][0],
            'PM OutTuesday': times_week1[2][1],
            'OutTuesday': times_week1[2][2],
            'InTuesday': times_week1[2][3],
            'AM InWednesday': times_week1[3][0],
            'PM OutWednesday': times_week1[3][1],
            'OutWednesday': times_week1[3][2],
            'InWednesday': times_week1[3][3],
            'AM InThursday': times_week1[4][0],
            'PM OutThursday': times_week1[4][1],
            'OutThursday': times_week1[4][2],
            'InThursday': times_week1[4][3],
            'AM InFriday': times_week1[5][0],
            'PM OutFriday': times_week1[5][1],
            'OutFriday': times_week1[5][2],
            'InFriday': times_week1[5][3],
            'Hours WorkedTotal for the Week': str(total_hours_week1),
            'DateSunday_2': dates_week2[0],
            'DateMonday_2': dates_week2[1],
            'DateTuesday_2': dates_week2[2],
            'DateWednesday_2': dates_week2[3],
            'DateThursday_2': dates_week2[4],
            'DateFriday_2': dates_week2[5],
            'DateSaturday_2': dates_week2[6],
            'Hours WorkedSunday_2': str(hours_week2[0]),
            'Hours WorkedMonday_2': str(hours_week2[1]),
            'Hours WorkedTuesday_2': str(hours_week2[2]),
            'Hours WorkedWednesday_2': str(hours_week2[3]),
            'Hours WorkedThursday_2': str(hours_week2[4]),
            'Hours WorkedFriday_2': str(hours_week2[5]),
            'Hours WorkedSaturday_2': str(hours_week2[6]),
            'AM InMonday_2': times_week2[1][0],
            'PM OutMonday_2': times_week2[1][1],
            'OutMonday_2': times_week2[1][2],
            'InMonday_2': times_week2[1][3],
            'AM InTuesday_2': times_week2[2][0],
            'PM OutTuesday_2': times_week2[2][1],
            'OutTuesday_2': times_week2[2][2],
            'InTuesday_2': times_week2[2][3],
            'AM InWednesday_2': times_week2[3][0],
            'PM OutWednesday_2': times_week2[3][1],
            'OutWednesday_2': times_week2[3][2],
            'InWednesday_2': times_week2[3][3],
            'AM InThursday_2': times_week2[4][0],
            'PM OutThursday_2': times_week2[4][1],
            'OutThursday_2': times_week2[4][2],
            'InThursday_2': times_week2[4][3],
            'AM InFriday_2': times_week2[5][0],
            'PM OutFriday_2': times_week2[5][1],
            'OutFriday_2': times_week2[5][2],
            'InFriday_2': times_week2[5][3],
            'Hours WorkedTotal for the Week_2': str(total_hours_week2),
            'Hours WorkedTotal for the Period': str(total_hours_period)
        }

        input_pdf_path = 'time-sheets.pdf'
        output_pdf_path = f'{pay_date}_{name}_timesheet_{pay_date}.pdf'
        fill_pdf(input_pdf_path, output_pdf_path, data)
