import pandas as pd

def process_data(input_file, output_file):
    # Read the data from RawData.txt
    df = pd.read_csv(input_file, delimiter="\t")

    # Group data by Employee Name, Report Date, and Activity Name
    grouped_data = df.groupby(['Employee Name', 'Report Date', 'Activity Name']).agg({
        'Activity Start': 'min',
        'Activity End': 'max',
        'Duration': 'sum'
    }).reset_index()

    # Format the duration to have two decimal places
    grouped_data['Duration'] = grouped_data['Duration'].map('{:.2f}'.format)

    # Write the processed data to Summaries.txt
    grouped_data.to_csv(output_file, sep='\t', index=False)

# Call the function with file names
process_data('RawData.txt', 'Summaries.txt')
