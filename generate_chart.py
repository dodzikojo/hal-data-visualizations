import pandas as pd
import plotly.express as px
import os
from datetime import datetime

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
            where = parts[2] if len(parts) > 2 else None
            return who, what, where
        
        df[['Who', 'What', 'Where']] = df['Created Roles'].apply(lambda x: pd.Series(parse_role(x)))
        df.dropna(subset=['Who', 'What', 'Created Roles'], inplace=True)
        df['count'] = 1
        
        # --- Extract Role Hierarchy ---
        # Build nested dictionary: {Level1: {Level2: [Level3 values]}}
        role_hierarchy = {}
        for _, row in df.iterrows():
            level1 = row['Who']
            level2 = row['What']
            level3 = row['Where'] if pd.notna(row['Where']) else 'All'
            
            if level1 not in role_hierarchy:
                role_hierarchy[level1] = {}
            if level2 not in role_hierarchy[level1]:
                role_hierarchy[level1][level2] = set()
            role_hierarchy[level1][level2].add(level3)
        
        # Convert sets to sorted lists for JSON serialization
        for level1 in role_hierarchy:
            for level2 in role_hierarchy[level1]:
                role_hierarchy[level1][level2] = sorted(list(role_hierarchy[level1][level2]))
        
        # Generate JavaScript object string
        import json
        hierarchy_json = json.dumps(role_hierarchy, indent=8, ensure_ascii=False)
        hierarchy_js = f"const roleHierarchy = {hierarchy_json}; // ROLE_HIERARCHY_PLACEHOLDER"
        
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
        
        # --- Inject Custom Script for Interactivity ---
        # We need to add a script to sends click events to the parent window
        custom_script = """
        <script>
            document.addEventListener("DOMContentLoaded", function() {
                var checkPlotly = setInterval(function() {
                    var plotDiv = document.getElementsByClassName('plotly-graph-div')[0];
                    if (plotDiv && plotDiv.data && plotDiv.data.length > 0) {
                        clearInterval(checkPlotly);
                        
                        // Send initial data to parent
                        var data = plotDiv.data[0];
                        window.parent.postMessage({
                            type: 'chartData',
                            data: {
                                ids: data.ids,
                                labels: data.labels,
                                parents: data.parents
                            }
                        }, '*');

                        // Handle Clicks
                        plotDiv.on('plotly_click', function(data) {
                            var point = data.points[0];
                            window.parent.postMessage({
                                type: 'chartClick',
                                data: {
                                    id: point.id,
                                    label: point.label,
                                    parent: point.parent,
                                    value: point.value
                                }
                            }, '*');
                        });

                        // Handle Double Clicks (Reset)
                        plotDiv.on('plotly_doubleclick', function() {
                            window.parent.postMessage({
                                type: 'chartDoubleClick'
                            }, '*');
                        });
                    }
                }, 500);
            });
        </script>
        """
        
        with open(chart_only_filepath, 'a', encoding='utf-8') as f:
            f.write(custom_script)

        print(f"Successfully generated chart file at: '{chart_only_filepath}'")
        
        # --- Update ACC_Roles_Report.html ---
        report_filepath = os.path.join(output_dir, "ACC_Roles_Report.html")
        if os.path.exists(report_filepath):
            try:
                with open(report_filepath, 'r', encoding='utf-8') as f:
                    report_html = f.read()
                
                # Generate timestamp
                generation_time = datetime.now().strftime("%B %d, %Y at %I:%M %p")
                
                # Replace the date in the span with id="chartGeneratedDate"
                import re
                updated_html = re.sub(
                    r'<span id="chartGeneratedDate">.*?</span>',
                    f'<span id="chartGeneratedDate">{generation_time}</span>',
                    report_html
                )
                
                # Replace the role hierarchy placeholder
                updated_html = re.sub(
                    r'const roleHierarchy = .*?; // ROLE_HIERARCHY_PLACEHOLDER',
                    hierarchy_js,
                    updated_html,
                    flags=re.DOTALL
                )
                
                with open(report_filepath, 'w', encoding='utf-8') as f:
                    f.write(updated_html)
                
                print(f"Updated generation timestamp in '{report_filepath}' to: {generation_time}")
                print(f"Embedded role hierarchy with {len(role_hierarchy)} Level 1 entries")
            except Exception as e:
                print(f"Warning: Could not update report: {e}")
        else:
            print(f"Warning: Report file not found at '{report_filepath}'")

    except FileNotFoundError:
        print(f"Error: Input file not found at '{excel_file_path}'")
    except Exception as e:
        print(f"An error occurred during chart generation: {e}")

if __name__ == "__main__":
    excel_filename = "HAL - ACC Roles - Consenting - Generated Roles.xlsx"
    generate_chart_html(excel_filename)