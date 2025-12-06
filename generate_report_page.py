import pandas as pd
import plotly.express as px
import os

def generate_static_report(excel_file_path):
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    
    # --- Define filenames ---
    chart_only_filename = "_chart_only.html"
    report_filename = "ACC_Roles_Report.html"
    
    chart_only_filepath = os.path.join(output_dir, chart_only_filename)
    report_filepath = os.path.join(output_dir, report_filename)

    # --- 1. Generate and save the standalone chart page ---
    try:
        df = pd.read_excel(excel_file_path)
        def parse_role(role_string):
            parts = [p.strip() for p in role_string.split(' - ')]
            who = parts[0] if len(parts) > 0 else None
            what = parts[1] if len(parts) > 1 else None
            return who, what
        df[['Who', 'What']] = df['Created Roles'].apply(lambda x: pd.Series(parse_role(x)))
        df.dropna(subset=['Who', 'What', 'Created Roles'], inplace=True)
        df['count'] = 1
        
        fig = px.sunburst(
            df,
            path=['Who', 'What', 'Created Roles'],
            values='count'
        )
        
        fig.update_traces(textinfo="label")
        fig.update_layout(
            margin=dict(t=80, l=40, r=40, b=40),
            title=dict(
                text="Sunburst Chart of Development Consent Order ACC Roles Distribution",
                x=0.05, xanchor='left', font=dict(size=22, color='#333')
            )
        )
        
        # Save the chart as a full, standalone HTML file. No JS post_script.
        fig.write_html(chart_only_filepath, full_html=True, include_plotlyjs='cdn')
        print(f"Standalone chart saved as '{chart_only_filepath}'")

    except Exception as e:
        print(f"Error generating chart: {e}")
        return

    # --- 2. Define the Explanatory Content for the main page ---
    explanation_html = """
        <p>This visualization isn't just a chart; it's a tool that helps us answer critical questions about project governance and security in ACC.</p>
        <h2>Why It's Useful: Answering Key Questions About Our Project</h2>
        <h3>It brings clarity to complexity.</h3>
        <p><strong>Question it answers:</strong> "How is our project <em>actually</em> structured?"<br>
        <strong>Explanation:</strong> This turns a complex list of hundreds of roles into an intuitive map. We can instantly see the entire role ecosystem in one place, which is invaluable for strategic planning and oversight.</p>
        <h3>It helps in managing resource allocation.</h3>
        <p><strong>Question it answers:</strong> "Are roles and permissions distributed appropriately across the different companies?"<br>
        <strong>Explanation:</strong> We can immediately spot imbalances. For example, we can see if one partner has a concentration of high-privilege roles, which helps us manage risk and ensure responsibilities are correctly aligned with contracts.</p>
        <h3>It's a powerful tool for security audits and onboarding.</h3>
        <p><strong>Question it answers:</strong> "Who has access to what, and is it correct?"<br>
        <strong>Explanation:</strong> When auditing our ACC setup, we can visually trace from a company to a specific role to its permissions. It's also perfect for onboarding new project leaders, as it gives them a rapid understanding of who the key players are and what access they have.</p>
    """

    # --- 3. Create the Final Report Page with an Iframe ---
    full_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>ACC Roles Distribution Report</title>
        <style>
            :root {{
                --primary-font: 'Segoe UI', 'Helvetica Neue', Helvetica, Arial, sans-serif;
                --background-color: #ffffff;
                --text-color: #555;
                --header-color: #222;
                --border-color: #e8e8e8;
            }}
            * {{ box-sizing: border-box; }}
            body {{
                font-family: var(--primary-font); margin: 0; background-color: var(--background-color);
                color: var(--text-color); line-height: 1.7; font-size: 16px;
                display: flex; min-height: 100vh;
            }}
            .main-content {{
                display: flex;
                flex-direction: column;
                width: 100%;
                padding: 2rem;
            }}
            .chart-container {{
                flex: 1;
                min-height: 600px;
            }}
            .chart-iframe {{
                width: 100%;
                height: 100%;
                border: none;
            }}
            .explanation-container {{
                margin-top: 2rem;
                border-top: 1px solid var(--border-color);
                padding-top: 2rem;
            }}
            h2, h3 {{
                color: var(--header-color);
                font-weight: 600;
                margin-top: 1rem;
                margin-bottom: 1rem;
            }}
            h2 {{ font-size: 1.8rem; }}
            h3 {{ font-size: 1.3rem; }}
            p {{ margin-bottom: 1.5rem; }}
            strong {{ color: var(--header-color); font-weight: 600; }}

            @media (min-width: 1200px) {{
                .main-content {{
                    flex-direction: row;
                }}
                .chart-container {{
                    flex-basis: 65%;
                    min-height: 90vh;
                }}
                .explanation-container {{
                    flex-basis: 35%;
                    padding-left: 3rem;
                    border-left: 1px solid var(--border-color);
                    border-top: none;
                    margin-top: 0;
                    padding-top: 0;
                }}
            }}
        </style>
    </head>
    <body>
        <main class="main-content">
            <section class="chart-container">
                <iframe class="chart-iframe" src="{chart_only_filename}"></iframe>
            </section>
            <aside class="explanation-container">
                {explanation_html}
            </aside>
        </main>
        <!-- No JavaScript needed for the static version -->
    </body>
    </html>
    """
    
    with open(report_filepath, 'w', encoding='utf-8') as f:
        f.write(full_html)
        
    print(f"Static report page has been regenerated and saved as '{report_filepath}'")

if __name__ == "__main__":
    excel_filename = "HAL - ACC Roles - Consenting - Generated Roles.xlsx"
    generate_static_report(excel_filename)
