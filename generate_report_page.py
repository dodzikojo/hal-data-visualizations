import pandas as pd
import plotly.express as px
import os

def generate_report_page(excel_file_path):
    # --- 1. Generate the Chart ---
    try:
        df = pd.read_excel(excel_file_path)
        # Parse data
        def parse_role(role_string):
            parts = [p.strip() for p in role_string.split(' - ')]
            who = parts[0] if len(parts) > 0 else 'N/A'
            what = parts[1] if len(parts) > 1 else 'N/A'
            return who, what
        df[['Who', 'What']] = df['Created Roles'].apply(lambda x: pd.Series(parse_role(x)))
        df.dropna(subset=['Who', 'What', 'Created Roles'], inplace=True)
        df['count'] = 1
        
        # Create the figure object
        fig = px.sunburst(
            df,
            path=['Who', 'What', 'Created Roles'],
            values='count',
            title="Sunburst Chart of Development Consent Order ACC Roles Distribution"
        )
        fig.update_traces(textinfo="label")
        fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))

        # Get the HTML representation of the chart
        chart_html = fig.to_html(full_html=False, include_plotlyjs='cdn')

    except Exception as e:
        print(f"Error generating chart: {e}")
        return

    # --- 2. Define the Explanatory Content ---
    
    explanation_html = """
        <h2>What It Is: The Big Picture of Our ACC Roles</h2>
        <p>
            What you're looking at is a <strong>sunburst chart</strong>, which is an interactive, hierarchical map of every single role we've defined for this project in Autodesk Construction Cloud. Instead of viewing these roles as a flat list in a spreadsheet, this chart visualizes the relationships between them, giving us a powerful, at-a-glance overview.
        </p>

        <h2>How to Read It: From the Companies to the Permissions</h2>
        <p>The chart is read from the inside out, telling a clear story about our project's structure.</p>
        <ul>
            <li>The <strong>inner ring</strong> represents the <strong>'Who'</strong>—the different companies and organizations involved in the project. The size of each segment immediately tells you which companies have the most roles assigned to them.</li>
            <li>As you move to the <strong>middle ring</strong>, you see the <strong>'What'</strong>. This breaks down each company into the specific functional roles they hold, like 'Approver', 'Designer', or 'Project Manager'. This shows us the makeup and primary function of each team.</li>
            <li>Finally, the <strong>outermost ring</strong> shows the <strong>full and exact role description</strong> as it exists in ACC. This is the most granular level, representing a specific set of permissions and responsibilities.</li>
        </ul>

        <h2>Why It's Useful: Answering Key Questions About Our Project</h2>
        <p>This visualization isn't just a chart; it's a tool that helps us answer critical questions about project governance and security in ACC.</p>
        
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

    # --- 3. Create the Final HTML Page ---
    
    full_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>ACC Roles Distribution Report</title>
        <style>
            body {{
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                margin: 0;
                background-color: #f8f9fa;
                color: #212529;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                padding: 20px;
            }}
            .header {{
                background-color: #fff;
                border-bottom: 1px solid #dee2e6;
                padding: 20px 40px;
                text-align: center;
            }}
            .header h1 {{
                margin: 0;
                font-size: 2.5rem;
            }}
            .content {{
                display: flex;
                flex-wrap: wrap;
                gap: 20px;
                margin-top: 20px;
            }}
            .chart-container {{
                flex: 2;
                min-width: 600px;
                background-color: #fff;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            .explanation-container {{
                flex: 1;
                min-width: 300px;
                background-color: #fff;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            h2, h3 {{
                border-bottom: 1px solid #eee;
                padding-bottom: 10px;
                margin-top: 20px;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>Autodesk Construction Cloud Roles Report</h1>
        </div>
        <div class="container">
            <div class="content">
                <div class="chart-container">
                    {chart_html}
                </div>
                <div class="explanation-container">
                    {explanation_html}
                </div>
            </div>
        </div>
    </body>
    </html>
    """

    # --- 4. Save the File ---
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    report_filename = os.path.join(output_dir, 'ACC_Roles_Report.html')
    
    with open(report_filename, 'w', encoding='utf-8') as f:
        f.write(full_html)
        
    print(f"Report page saved as '{report_filename}'")


if __name__ == "__main__":
    excel_filename = "HAL - ACC Roles - Consenting - Generated Roles.xlsx"
    generate_report_page(excel_filename)
