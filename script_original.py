import pandas as pd
from datetime import datetime
import os

def process_file(filename):
    # Specify column names
    column_names = ['Date_Time', 'Data_Type', 'Process', 'Fat', 'SNF', 'Density', 'Water', 'LTC', 'GSS', 'Factory']

    # Read the TXT file into a pandas DataFrame
    data = pd.read_csv(filename, sep='\t', header=None, names=column_names, na_values=[''])

    # Convert 'Date_Time' column to strings
    data['Date_Time'] = data['Date_Time'].astype(str)

    # Function to convert the date-time string to a datetime object
    def convert_to_datetime(number):
        try:
            if len(number) < 12:
                number = '0' + number

            day = int(number[0:2])
            month = int(number[2:4])
            year = 2000 + int(number[4:6])
            hour = int(number[6:8])
            minute = int(number[8:10])
            second = int(number[10:12])

            # Check for valid date and time
            if month < 1 or month > 12 or day < 1 or day > 31 or hour > 23 or minute > 59 or second > 59:
                return None

            # Attempt to create a datetime object to catch invalid dates (e.g., February 30)
            return datetime(year, month, day, hour, minute, second)
        except (ValueError, TypeError):
            return None

    # Apply the function to the 'Date_Time' column
    data['Date_Time'] = data['Date_Time'].apply(convert_to_datetime)

    # Remove rows with invalid date-time values
    data = data.dropna(subset=['Date_Time'])

    # Custom function to determine shift
    def get_shift(datetime_value):
        hour = datetime_value.hour
        if 0 < hour < 12:
            return 'Morning'
        else:
            return 'Evening'

    # Apply the function to create 'Shift' column
    data['Shift'] = data['Date_Time'].apply(get_shift)

    data['Date'] = data['Date_Time'].dt.date
    data['Year'] = data['Date_Time'].dt.year
    data['Month'] = data['Date_Time'].dt.month
    data['Day'] = data['Date_Time'].dt.day
    data['Month_Name'] = data['Date_Time'].dt.month_name()

    # Calculate number of unique dates
    number_of_records = len(data['Date'].unique())
    One = pd.DataFrame({'Number_of_Records': [number_of_records]})
    
    # Processing Two: Pivot table of MS process by Date and Shift
    Two = data[data['Process'] == 'MS'].pivot_table(index='Date', columns='Shift', aggfunc={'Process': 'size'}).sort_index(ascending=False)

    # Processing Three: Sum of Process1 by Date and Shift divided by 25
    data['Process1'] = pd.to_numeric(data['Process'], errors='coerce')
    Three = round(data.pivot_table(index='Date', columns='Shift', aggfunc={'Process1': 'sum'}) / 25, 0).sort_index(ascending=False)


    # Save the processed DataFrame and One DataFrame to an output Excel file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f'processed_{timestamp}.xlsx'
    output_filepath = os.path.join('output', output_filename)
    
    with pd.ExcelWriter(output_filepath) as writer:
        data.to_excel(writer, sheet_name='Processed_Data', index=False)
        One.to_excel(writer, sheet_name='One', index=False)
        Two.to_excel(writer, sheet_name='Two')
        Three.to_excel(writer, sheet_name='Three')

    return output_filename

# Example usage:
if __name__ == "__main__":
    input_filepath = 'path_to_your_input_file.txt'  # Replace with your input file path
    result_file = process_file(input_filepath)
    print(f"Processed data saved to: {result_file}")
