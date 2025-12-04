import pandas as pd
import plotly.express as px
import os

def create_sunburst_custom_labels(excel_file_path):
    # Check if the file exists
    if not os.path.exists(excel_file_path):
        print(f"Error: The file '{excel_file_path}' was not found.")
        return

    try:
        df = pd.read_excel(excel_file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    if 'Created Roles' not in df.columns:
        print("Error: 'Created Roles' column not found.")
        return

    # Parse the 'Created Roles' column into Who and What
    def parse_role(role_string):
        parts = [p.strip() for p in role_string.split(' - ')]
        who = parts[0] if len(parts) > 0 else 'N/A'
        what = parts[1] if len(parts) > 1 else 'N/A'
        return who, what

    df[['Who', 'What']] = df['Created Roles'].apply(lambda x: pd.Series(parse_role(x)))
    df.dropna(subset=['Who', 'What', 'Created Roles'], inplace=True)


    # The sunburst chart needs a 'values' column. We'll add a count of 1 for each row.
    df['count'] = 1

    # Create the sunburst chart with the new path and title
    fig = px.sunburst(
        df,
        path=['Who', 'What', 'Created Roles'], # New path for custom labels
        values='count',
        title="Sunburst Chart of Development Consent Order ACC Roles Distribution" # New title
    )
    # Remove percentages from text labels
    fig.update_traces(textinfo="label")
    fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))

    # --- Saving ---
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    base_filename = os.path.splitext(os.path.basename(excel_file_path))[0]
    # New filename for the modified chart
    chart_filename = os.path.join(output_dir, f'{base_filename}_sunburst_custom_labels.html')

    fig.write_html(chart_filename)
    print(f"Sunburst chart with custom labels saved as '{chart_filename}'")

if __name__ == "__main__":
    excel_filename = "HAL - ACC Roles - Consenting - Generated Roles.xlsx"
    create_sunburst_custom_labels(excel_filename)
