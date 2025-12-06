import pandas as pd
import plotly.express as px
import os

def generate_chart_html(excel_file_path):
    """
    Reads an Excel file, processes the role data, and generates a standalone
    HTML file for the sunburst chart.
    """
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    
    chart_only_filename = "_chart_only.html"
    chart_only_filepath = os.path.join(output_dir, chart_only_filename)

    try:
        df = pd.read_excel(excel_file_path)
        
        # --- Data Processing ---
        def parse_role(role_string):
            parts = [p.strip() for p in role_string.split(' - ')]
            who = parts[0] if len(parts) > 0 else None
            what = parts[1] if len(parts) > 1 else None
            return who, what
        df[['Who', 'What']] = df['Created Roles'].apply(lambda x: pd.Series(parse_role(x)))
        df.dropna(subset=['Who', 'What', 'Created Roles'], inplace=True)
        df['count'] = 1
        
        # --- Figure Creation ---
        fig = px.sunburst(df, path=['Who', 'What', 'Created Roles'], values='count')
        
        fig.update_traces(textinfo="label")
        fig.update_layout(
            margin=dict(t=80, l=40, r=40, b=40),
            title=dict(
                text="Sunburst Chart of Development Consent Order ACC Roles Distribution",
                x=0.05, xanchor='left', font=dict(size=22, color='#333')
            )
        )
        
        # --- HTML Generation ---
        fig.write_html(chart_only_filepath, full_html=True, include_plotlyjs='cdn')
        print(f"Successfully generated chart file at: '{chart_only_filepath}'")

    except FileNotFoundError:
        print(f"Error: Input file not found at '{excel_file_path}'")
    except Exception as e:
        print(f"An error occurred during chart generation: {e}")

if __name__ == "__main__":
    excel_filename = "HAL - ACC Roles - Consenting - Generated Roles.xlsx"
    generate_chart_html(excel_filename)