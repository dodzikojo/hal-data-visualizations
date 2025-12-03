import pandas as pd
import matplotlib.pyplot as plt
import os

def create_role_frequency_chart(excel_file_path):
    # Check if the file exists
    if not os.path.exists(excel_file_path):
        print(f"Error: The file '{excel_file_path}' was not found.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(excel_file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Check if the 'Created Roles' column exists
    if 'Created Roles' not in df.columns:
        print("Error: 'Created Roles' column not found in the Excel file.")
        return

    # Extract the job title from the 'Created Roles' string
    # Assumes the format "Code - Title - Location"
    def extract_title(role_string):
        parts = role_string.split(' - ')
        if len(parts) > 1:
            return parts[1].strip()
        return None

    df['Job Title'] = df['Created Roles'].apply(extract_title)

    # Count the frequency of each job title
    title_counts = df['Job Title'].value_counts().dropna()

    if title_counts.empty:
        print("Could not extract any job titles. Please check the data format.")
        return

    # Create the horizontal bar chart
    plt.figure(figsize=(12, 8))
    title_counts.sort_values().plot(kind='barh')
    plt.title('Frequency of Job Titles')
    plt.xlabel('Count')
    plt.ylabel('Job Title')
    plt.grid(axis='x')
    plt.tight_layout() # Adjust layout to make room for labels

    # Create output directory if it doesn't exist
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)

    # Save the chart with high resolution
    base_filename = os.path.splitext(os.path.basename(excel_file_path))[0]
    chart_filename = os.path.join(output_dir, f'{base_filename}_role_frequency.png')
    plt.savefig(chart_filename, dpi=300)
    print(f"Bar chart saved as '{chart_filename}'")
    plt.show() # Display the plot

if __name__ == "__main__":
    excel_filename = "HAL - ACC Roles - Consenting - Generated Roles.xlsx"
    create_role_frequency_chart(excel_filename)
