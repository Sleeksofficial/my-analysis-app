import pandas as pd
import io

# Function to read and analyze Excel file
def analyze_excel(file_path_or_buffer):
    """
    Reads an Excel file (from a path or an in-memory buffer) 
    and returns a summary, a dictionary of DataFrames, and an error message (if any).
    """
    try:
        data = pd.read_excel(file_path_or_buffer, sheet_name=None)  # Load all sheets
        sheets = data.keys()

        # --- THIS IS THE FIX ---
        if not sheets:
            return None, None, "The uploaded Excel file contains no sheets or is empty."
        # ----------------------

        summary = {}
        for sheet_name in sheets:
            df = data[sheet_name]
            summary[sheet_name] = {
                'Shape': df.shape,
                'Columns': df.columns.tolist(),
                'Numeric_Columns': df.select_dtypes(include='number').columns.tolist(),
                'Categorical_Columns': df.select_dtypes(exclude='number').columns.tolist(),
                'Head': df.head().to_dict(orient='records')
            }
        
        # Return None for the error
        return summary, data, None
        
    except Exception as e:
        print(f"Error loading the Excel file: {e}")
        # Return the error message
        return None, None, str(e)
# --- UPGRADED "AI" REPORTING FUNCTION ---
def generate_report(summary, data):
    report = "## 📊 Automated Data Analysis Report\n\n"
    
    if not summary:
        report += "### ⚠️ Error\nNo data could be loaded for analysis."
        return report

    report += f"The uploaded workbook contains **{len(summary)} sheet(s)**.\n\n"
    
    time_keywords = ['date', 'day', 'month', 'year', 'timestamp', 'time']
    cat_keywords = ['region', 'country', 'city', 'state', 'department', 'category', 'gender', 'status', 'type', 'group']
    measure_keywords = ['sales', 'amount', 'revenue', 'count', 'cases', 'salary', 'price', 'quantity', 'value', 'score', 'rate', 'cost', 'profit']
    id_keywords = ['id', 'uuid', 'key', 'code', 'number']

    for sheet_name, sheet_summary in summary.items():
        df = data[sheet_name]
        rows, cols = sheet_summary['Shape']
        num_cols_list = sheet_summary['Numeric_Columns']
        cat_cols_list = sheet_summary['Categorical_Columns']
        
        report += f"### Sheet Analysis: `{sheet_name}`\n\n"
        
        # --- 1. Structural Summary ---
        report += f"#### 1. Structural Summary\n"
        report += f"* This sheet has **{rows} rows** and **{cols} columns**.\n"
        report += f"* It contains **{len(num_cols_list)} numeric columns** and **{len(cat_cols_list)} categorical columns**.\n\n"

        # --- 2. Thematic Summary (Heuristic) ---
        report += f"#### 2. Thematic Summary & Advice\n\n"
        
        time_cols = [c for c in df.columns if any(kw in c.lower() for kw in time_keywords)]
        measure_cols = [c for c in num_cols_list if any(kw in c.lower() for kw in measure_keywords)]
        cat_cols = [c for c in cat_cols_list if any(kw in c.lower() for kw in cat_keywords)]
        id_cols = [c for c in cat_cols_list if any(kw in c.lower() for kw in id_keywords)]

        report += "**What is this file? (Inferred Summary):**\n"
        if time_cols and measure_cols and cat_cols:
            report += f"* This sheet appears to be **time-series data for monitoring metrics**. \n"
            report += f"* It likely tracks key figures (like `{measure_cols[0]}`) over time (using `{time_cols[0]}`) and segments them by categories (like `{cat_cols[0]}`).\n"
        elif measure_cols and cat_cols:
            report += f"* This sheet appears to be **transactional or observational data**.\n"
            report += f"* It measures key metrics (like `{measure_cols[0]}`) across different categories (like `{cat_cols[0]}`).\n"
        else:
            report += f"* The purpose of this file is general. It contains various numeric and categorical fields.\n"

        report += "\n**How to analyze this? (Analysis Advice):**\n"
        if time_cols:
            report += f"* **Use Line Charts** to see trends. Plot a metric on the Y-Axis against your time column (`{time_cols[0]}`) on the X-Axis.\n"
        if cat_cols:
            report += f"* **Use Bar Charts** to compare groups. Plot a metric on the Y-Axis against a category (like `{cat_cols[0]}`) on the X-Axis.\n"
        if len(measure_cols) >= 2:
            report += f"* **Use Scatter Plots** to find relationships. Plot one metric (like `{measure_cols[0]}`) on the X-Axis and another (like `{measure_cols[1]}`) on the Y-Axis.\n"
        if time_cols:
            report += f"* **Use the 'Forecasting' Tab** to predict future values. Select `{time_cols[0]}` as your date column.\n"

        # --- 3. Data Criticism & Quality Issues ---
        report += f"\n#### 3. Data Criticism & Quality Issues\n\n"
        
        missing_total = df.isnull().sum().sum()
        duplicate_count = df.duplicated().sum()
        
        if missing_total == 0 and duplicate_count == 0:
            report += "* ✅ **Excellent!** No missing values or duplicate rows were found. This data is clean.\n"
        
        if missing_total > 0:
            report += f"* ⚠️ **Missing Data:** This sheet has **{missing_total} missing values**.\n"
            report += f"    * **Criticism:** Missing data can skew averages, break charts, and cause AI models to fail. Use the 'Data Cleaning' tab to drop these rows.\n"
                
        if duplicate_count > 0:
            report += f"* ⚠️ **Duplicate Rows:** **{duplicate_count} identical rows** were found.\n"
            report += f"    * **Criticism:** This will lead to double-counting and inflated totals. Use the 'Data Cleaning' tab to remove these.\n"
            
        # --- 4. NEW: Key Insights & Discoveries ---
        report += f"\n#### 4. Key Insights & Discoveries (AI-Generated)\n\n"
        
        if not num_cols_list:
             report += "* No numeric data found to generate insights.\n"
        else:
            # Insight 1: Correlations
            try:
                corr_matrix = df[num_cols_list].corr()
                # Find the strongest positive/negative correlations
                corr_pairs = corr_matrix.unstack().sort_values(ascending=False)
                # Remove self-correlations
                corr_pairs = corr_pairs[corr_pairs != 1.0]
                
                if not corr_pairs.empty:
                    strongest_pos = corr_pairs.head(1).index[0]
                    strongest_neg = corr_pairs.tail(1).index[0]
                    
                    if corr_pairs.max() > 0.7:
                        report += f"* **Strong Positive Correlation:** There is a strong relationship (`{corr_pairs.max():.2f}`) between `{strongest_pos[0]}` and `{strongest_pos[1]}`. When one goes up, the other tends to go up as well.\n"
                    if corr_pairs.min() < -0.7:
                        report += f"* **Strong Negative Correlation:** There is a strong inverse relationship (`{corr_pairs.min():.2f}`) between `{strongest_neg[0]}` and `{strongest_neg[1]}`. When one goes up, the other tends to go down.\n"
                else:
                    report += "* No significant correlations were found between numeric columns.\n"
            except Exception as e:
                report += f"* Could not calculate correlations: {e}\n"

            # Insight 2: Low Performers
            profit_cols = [c for c in measure_cols if 'profit' in c.lower()]
            if profit_cols and (df[profit_cols[0]].mean() < 0):
                report += f"* **Financial Warning:** The average for `{profit_cols[0]}` is **negative**. The business may be losing money on average.\n"
            elif 'sales' in measure_cols and df['sales'].mean() < 100:
                 report += f"* **Performance Insight:** The average `sales` value is very low. Use the dashboard to investigate which categories or regions are underperforming.\n"
            else:
                report += "* All numeric metrics appear to be within a standard positive range.\n"

        report += "\n---\n"
    return report