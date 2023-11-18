import pandas as pd

def process_emails(input_file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(input_file_path)

        # Remove duplicates from 'master email' column
        df['master email'] = df['master email'].drop_duplicates()

        # Remove duplicates from 'new email' column
        df['new email'] = df['new email'].drop_duplicates()
        #new_emails_no_duplicates = df['new email']

        # Create a new column with emails without duplicates
        df['no duplicates email'] = df['new email'].drop_duplicates()
        new_emails_no_duplicates = df['no duplicates email']

        # Concatenate 'master email' and non-duplicate new emails
        df['master email'] = pd.concat([pd.DataFrame(df['master email']), pd.DataFrame({'master email': new_emails_no_duplicates})], ignore_index=True).drop_duplicates()

        # Remove duplicates from 'master email' column after concatenation
        df['master email'] = df['master email'].drop_duplicates()

        # Create a new column with emails without duplicates
        #df['no duplicates email'] = df['new email'].drop_duplicates()

        # Sort 'New eMails Duplicates Removed' column by domain
        def extract_domain(email):
            parts = str(email).split('@')
            return parts[1] if len(parts) > 1 else None

        df = df.sort_values(by='no duplicates email', key=lambda x: x.apply(extract_domain), na_position='first').dropna()

        # Create a 'Domain' column
        df['Domain'] = df['no duplicates email'].apply(extract_domain)

        # Save the updated list to the input file (overwrite the original file)
        df.to_excel(input_file_path, index=False)
        print("Processing complete.")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    input_file_path = input("Enter the path of your Excel file: ").strip()

    if input_file_path:
        process_emails(input_file_path)
    else:
        print("No file path provided.")
