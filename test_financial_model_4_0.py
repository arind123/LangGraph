import pandas as pd
import numpy as np

class ExcelModel:
    def __init__(self):
        pass

    def calculate_price_adjusted(self, sales_df, products_df, discounts_df, tax_df):
        """
        Excel Formula: 
        =VLOOKUP(B2,Products!$A$2:$C$5,3, FALSE) * C2 - 
        (C2 * VLOOKUP(B2,Products!$A$2:$C$5, 3, FALSE) * VLOOKUP(B2,Tax!$A$2:$B$5, 2, FALSE)) - 
        (C2 * VLOOKUP(B2, Discounts!$A$2:$B$5, 2, FALSE))
        
        Dependencies: 
        - Products.BasePrice
        - Products.ProductID
        - Discounts.Product
        - Sales.Quantity
        - Discounts.Discount
        - Tax.Products
        - Sales.ProductID
        - Tax.Tax_Rate
        """
        sales_df = sales_df.copy()
        sales_df = sales_df.merge(products_df[['ProductID', 'BasePrice']], on='ProductID', how='left')
        sales_df = sales_df.merge(discounts_df, left_on='ProductID', right_on='Product', how='left')
        sales_df = sales_df.merge(tax_df, left_on='ProductID', right_on='Products', how='left')

        sales_df['Discount'] = sales_df['Discount'].fillna(0)
        sales_df['Tax_Rate'] = sales_df['Tax_Rate'].fillna(0)

        sales_df['Price_Adjusted'] = (
            sales_df['BasePrice'] * sales_df['Quantity'] -
            (sales_df['Quantity'] * sales_df['BasePrice'] * sales_df['Tax_Rate']) -
            (sales_df['Quantity'] * sales_df['Discount'])
        )
        return sales_df

    def calculate_total_sales(self, sales_df):
        """
        Excel Formula: 
        =SUM(Sales!D2:D6)
        
        Dependencies: 
        - Sales.Price_Adjusted
        """
        total_sales = sales_df['Price_Adjusted'].sum()
        return total_sales

    def calculate_global_margins_value(self, total_sales, global_margins_df):
        """
        Excel Formula: 
        =A2-(A2*Global_Margins!B2)
        
        Dependencies: 
        - Financials.Total_Sales
        - Global_Margins.Value
        """
        global_margins_value = total_sales - (total_sales * global_margins_df['Value'].iloc[0])
        return global_margins_value

    def calculate_profit_loss(self, global_margins_value, global_margins_df):
        """
        Excel Formula: 
        =IF(B2>Global_Margins!A2, "Profit", "Loss")
        
        Dependencies: 
        - Financials.Global_Margins_Value
        - Global_Margins.Global_Margin
        """
        profit_loss = "Profit" if global_margins_value > global_margins_df['Global_Margin'].iloc[0] else "Loss"
        return profit_loss

    def transform(self, all_sheets_dict):
        sales_df = all_sheets_dict['Sales']
        products_df = all_sheets_dict['Products']
        discounts_df = all_sheets_dict['Discounts']
        tax_df = all_sheets_dict['Tax']
        global_margins_df = all_sheets_dict['Global_Margins']

        # Calculate Price Adjusted
        sales_df = self.calculate_price_adjusted(sales_df, products_df, discounts_df, tax_df)

        # Calculate Total Sales
        total_sales = self.calculate_total_sales(sales_df)

        # Calculate Global Margins Value
        global_margins_value = self.calculate_global_margins_value(total_sales, global_margins_df)

        # Calculate Profit or Loss
        profit_loss = self.calculate_profit_loss(global_margins_value, global_margins_df)

        # Create Financials DataFrame
        financials_df = pd.DataFrame({
            'Total_Sales': [total_sales],
            'Global_Margins_Value': [global_margins_value],
            'Profit_Loss': [profit_loss]
        })

        return {
            'Sales': sales_df,
            'Financials': financials_df
        }