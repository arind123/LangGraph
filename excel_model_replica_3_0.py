import pandas as pd
import numpy as np

class ExcelModel:
    def __init__(self):
        pass

    def calculate_price_adjusted(self, sales_df, products_df, tax_df, discounts_df):
        """
        Excel Formula: =VLOOKUP(B2,Products!$A$2:$C$5,3, FALSE) * C2 - (C2 * VLOOKUP(B2,Products!$A$2:$C$5, 3, FALSE) * VLOOKUP(B2,Tax!$A$2:$B$5, 2, FALSE)) - (C2 * VLOOKUP(B2, Discounts!$A$2:$B$5, 2, FALSE))
        Dependencies: ['Tax.Products', 'Products.ProductID', 'Sales.Quantity', 'Discounts.Discount', 'Tax.Tax_Rate', 'Products.BasePrice', 'Discounts.Product', 'Sales.ProductID']
        Target Type: int
        """
        sales_df = sales_df.copy()
        sales_df = sales_df.merge(products_df[['ProductID', 'BasePrice']], left_on='ProductID', right_on='ProductID', how='left')
        sales_df = sales_df.merge(tax_df[['Products', 'Tax_Rate']], left_on='ProductID', right_on='Products', how='left')
        sales_df = sales_df.merge(discounts_df[['Product', 'Discount']], left_on='ProductID', right_on='Product', how='left')

        sales_df['Price_Adjusted'] = (
            sales_df['BasePrice'] * sales_df['Quantity'] -
            (sales_df['Quantity'] * sales_df['BasePrice'] * sales_df['Tax_Rate']) -
            (sales_df['Quantity'] * sales_df['Discount'])
        ).astype(int)

        return sales_df

    def calculate_total_sales(self, sales_df):
        """
        Excel Formula: =SUM(Sales!D2:D6)
        Dependencies: ['Sales.Price_Adjusted']
        Target Type: float
        """
        total_sales = sales_df['Price_Adjusted'].sum()
        return total_sales

    def calculate_global_margins_value(self, total_sales, global_margins_df):
        """
        Excel Formula: =A2-(A2*Global_Margins!B2)
        Dependencies: ['Global_Margins.Value', 'Financials.Total_Sales']
        Target Type: float
        """
        global_margins_value = total_sales - (total_sales * global_margins_df['Value'].iloc[0])
        return global_margins_value

    def calculate_profit_loss(self, global_margins_value, global_margins_df):
        """
        Excel Formula: =IF(B2>Global_Margins!A2, "Profit", "Loss")
        Dependencies: ['Global_Margins.Global_Margin', 'Financials.Global_Margins_Value']
        Target Type: str
        """
        profit_loss = np.where(global_margins_value > global_margins_df['Global_Margin'].iloc[0], "Profit", "Loss")
        return profit_loss

    def transform(self, all_sheets_dict):
        # Extract dataframes from the dictionary
        sales_df = all_sheets_dict['Sales']
        products_df = all_sheets_dict['Products']
        tax_df = all_sheets_dict['Tax']
        discounts_df = all_sheets_dict['Discounts']
        global_margins_df = all_sheets_dict['Global_Margins']

        # Calculate Price Adjusted
        sales_df = self.calculate_price_adjusted(sales_df, products_df, tax_df, discounts_df)

        # Calculate Total Sales
        total_sales = self.calculate_total_sales(sales_df)

        # Calculate Global Margins Value
        global_margins_value = self.calculate_global_margins_value(total_sales, global_margins_df)

        # Calculate Profit or Loss
        profit_loss = self.calculate_profit_loss(global_margins_value, global_margins_df)

        # Add calculated columns to the Financials sheet
        financials_df = pd.DataFrame({
            'Total_Sales': [total_sales],
            'Global_Margins_Value': [global_margins_value],
            'Profit_Loss': [profit_loss]
        })

        # Update the all_sheets_dict with the new Financials dataframe
        all_sheets_dict['Financials'] = financials_df

        return all_sheets_dict