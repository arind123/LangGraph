import pandas as pd
import numpy as np

class ExcelModel:
    def __init__(self):
        pass

    def calculate_value_1(self, task_tracker_df, dept_category):
        """
        Excel Formula: =SUMIFS(Task_Tracker!G:G, Task_Tracker!B:B, A2)
        Dependencies: Budget_Analysis.Dept_Category, Task_Tracker.Department
        """
        return task_tracker_df.loc[task_tracker_df['Department'] == dept_category, 'Hours_Logged'].sum()

    def calculate_value_2(self, value_1, allocated_budget):
        """
        Excel Formula: =IFERROR(C2/B2, 0) and =B2-C2
        Dependencies: Budget_Analysis.Value_1, Budget_Analysis.Allocated_Budget
        """
        with np.errstate(divide='ignore', invalid='ignore'):
            value_2_ratio = np.divide(value_1, allocated_budget, out=np.zeros_like(value_1, dtype=float), where=allocated_budget!=0)
        value_2_difference = allocated_budget - value_1
        return value_2_ratio, value_2_difference

    def calculate_value(self, budget_analysis_df):
        """
        Excel Formula: =SUM(Budget_Analysis!C:C)
        Dependencies: Budget_Analysis.Value_1
        """
        return budget_analysis_df['Value_1'].sum()

    def transform(self, all_sheets_dict):
        # Extracting Task_Tracker DataFrame
        task_tracker_df = all_sheets_dict.get('Task_Tracker', pd.DataFrame(columns=['Task_ID', 'Department', 'Resource_Name', 'Hourly_Rate', 'Hours_Logged', 'Status']))

        # Extracting Budget_Analysis DataFrame
        budget_analysis_df = all_sheets_dict.get('Budget_Analysis', pd.DataFrame(columns=['Dept_Category', 'Allocated_Budget']))

        # Calculate Value_1 for Budget_Analysis
        budget_analysis_df['Value_1'] = budget_analysis_df['Dept_Category'].apply(lambda dept: self.calculate_value_1(task_tracker_df, dept))

        # Calculate Value_2 for Budget_Analysis
        budget_analysis_df['Value_2_Ratio'], budget_analysis_df['Value_2_Difference'] = self.calculate_value_2(budget_analysis_df['Value_1'], budget_analysis_df['Allocated_Budget'])

        # Initialize Executive_Dashboard DataFrame
        executive_dashboard_df = pd.DataFrame(columns=['Metric', 'Value'])

        # Calculate Value for Executive_Dashboard
        executive_dashboard_value = self.calculate_value(budget_analysis_df)
        executive_dashboard_df = pd.concat([executive_dashboard_df, pd.DataFrame({'Metric': ['Total Value'], 'Value': [executive_dashboard_value]})], ignore_index=True)

        # Return the transformed DataFrames
        return {
            'Task_Tracker': task_tracker_df,
            'Budget_Analysis': budget_analysis_df,
            'Executive_Dashboard': executive_dashboard_df
        }