import pandas as pd

def refine_essraw(essraw_file, output_file):
    # Read the data from ESSRAW.txt
    essraw = pd.read_csv(essraw_file, delimiter="\t")

    # Function to reformat names
    def reformat_name(name):
        parts = name.split(' ')
        if len(parts) >= 2:
            # Remove extra spaces and format as 'LastName, FirstName'
            return parts[-1] + ', ' + ' '.join(parts[:-1]).strip()
        else:
            return name  # Fallback for names that don't split correctly

    # Apply the reformatting function to the 'Initiator' column
    essraw['Initiator'] = essraw['Initiator'].apply(reformat_name)

    # Select the necessary columns
    ess_summary = essraw[['Submit Date', 'Initiator', 'Approver', 'Att./abs. type text', 'Start Date', 'End Date', 'Absence hrs']]

    # Write to output file
    ess_summary.to_csv(output_file, sep='\t', index=False)

# Call the function with file names
refine_essraw('ESSRAW.txt', 'ESSummaries.txt')
