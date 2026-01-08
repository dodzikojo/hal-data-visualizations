import pandas as pd
import plotly.graph_objects as go
import os

def create_sankey_who_where(excel_file_path):
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

    # Parse the 'Created Roles' column
    def parse_role(role_string):
        parts = [p.strip() for p in role_string.split(' - ')]
        who = parts[0] if len(parts) > 0 else None
        # what = parts[1] if len(parts) > 1 else None # 'What' is no longer needed for the links
        where = parts[2] if len(parts) > 2 else None
        return who, where

    df[['Who', 'Where']] = df['Created Roles'].apply(lambda x: pd.Series(parse_role(x)))
    # Drop rows where the essential links are missing
    df.dropna(subset=['Who', 'Where'], inplace=True)
    
    if df.empty:
        print("Could not parse enough data (Who, Where) to create a Sankey diagram.")
        return

    # --- Prepare data for Sankey ---
    # Create a list of all unique labels (nodes) from Who and Where
    all_labels = pd.unique(df[['Who', 'Where']].values.ravel('K')).tolist()
    
    # Create a mapping from label to a unique integer
    label_map = {label: i for i, label in enumerate(all_labels)}
    
    # Map the dataframe columns to integers
    df_mapped = df.replace(label_map)

    # --- Create links from Who to Where ---
    links_who_where = df_mapped.groupby(['Who', 'Where']).size().reset_index(name='value')

    # --- Create the Sankey Diagram ---
    fig = go.Figure(data=[go.Sankey(
        node=dict(
            pad=15,
            thickness=20,
            line=dict(color="black", width=0.5),
            label=all_labels
        ),
        link=dict(
            source=links_who_where['Who'].tolist(),
            target=links_who_where['Where'].tolist(),
            value=links_who_where['value'].tolist()
        )
    )])

    fig.update_layout(title_text="Sankey Diagram of Connections from 'Who' to 'Where'", font_size=10)

    # --- Saving ---
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    base_filename = os.path.splitext(os.path.basename(excel_file_path))[0]
    chart_filename = os.path.join(output_dir, f'{base_filename}_sankey_who_where.html')

    fig.write_html(chart_filename)
    print(f"Modified Sankey diagram saved as '{chart_filename}'")


if __name__ == "__main__":
    excel_filename = "HAL - ACC Roles - Consenting - Generated Roles.xlsx"
    create_sankey_who_where(excel_filename)
