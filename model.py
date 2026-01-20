import pandas as pd
import numpy as np
from typing import Dict

class ExcelCalculator:
    def __init__(self, df: pd.DataFrame, all_data: Dict[str, pd.DataFrame]):
        self.df = df
        self.all_data = all_data

    @property
    def price_adjusted(self):
        products_df = self.all_data['Products']
        tax_df = self.all_data['Tax']
        discounts_df = self.all_data['Discounts']

        # Merge dataframes to get necessary columns
        merged_df = self.df.merge(products_df[['ProductID', 'BasePrice']], on='ProductID', how='left')
        merged_df = merged_df.merge(tax_df[['Products', 'Tax_Rate']], left_on='ProductID', right_on='Products', how='left')
        merged_df = merged_df.merge(discounts_df[['Product', 'Discount']], left_on='ProductID', right_on='Product', how='left')

        # Calculate Price_Adjusted
        base_price = merged_df['BasePrice']
        quantity = merged_df['Quantity']
        tax_rate = merged_df['Tax_Rate'].fillna(0)
        discount = merged_df['Discount'].fillna(0)

        price_adjusted = (base_price * quantity) - (quantity * base_price * tax_rate) - (quantity * discount)
        return price_adjusted

    @property
    def total_sales(self):
        sales_df = self.all_data['Sales']
        return sales_df['Price_Adjusted'].sum()

    @property
    def global_margins(self):
        global_margins_df = self.all_data['Global_Margins']
        total_sales = self.total_sales
        global_margin_value = global_margins_df['Value'].iloc[0]  # Assuming single value
        return total_sales - (total_sales * global_margin_value)

    @property
    def profit_loss(self):
        global_margins_df = self.all_data['Global_Margins']
        global_margin_threshold = global_margins_df['Global_Margin'].iloc[0]  # Assuming single value
        global_margins = self.global_margins
        return np.where(global_margins > global_margin_threshold, "Profit", "Loss")