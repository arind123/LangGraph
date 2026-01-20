import pandas as pd
import numpy as np

class ExcelModel:
    def __init__(self, products_df, tax_df, discounts_df, global_margins_df):
        self.products_df = products_df
        self.tax_df = tax_df
        self.discounts_df = discounts_df
        self.global_margins_df = global_margins_df

    def calculate_price_adjusted(self, df):
        '''
        Excel Formula: =VLOOKUP(B2,Products!$A$2:$C$5,3, FALSE) * C2 - (C2 * VLOOKUP(B2,Products!$A$2:$C$5, 3, FALSE) * VLOOKUP(B2,Tax!$A$2:$B$5, 2, FALSE)) - (C2 * VLOOKUP(B2, Discounts!$A$2:$B$5, 2, FALSE))
        Dependencies: Tax.Products, Discounts.Discount, Products.ProductID, Discounts.Product, Sales.ProductID, Tax.Tax_Rate, Products.BasePrice, Sales.Quantity
        '''
        # Merge Sales with Products to get BasePrice
        df = df.merge(self.products_df[['ProductID', 'BasePrice']], left_on='ProductID', right_on='ProductID', how='left')

        # Merge Sales with Tax to get Tax_Rate
        df = df.merge(self.tax_df[['Products', 'Tax_Rate']], left_on='ProductID', right_on='Products', how='left')

        # Merge Sales with Discounts to get Discount
        df = df.merge(self.discounts_df[['Product', 'Discount']], left_on='ProductID', right_on='Product', how='left')

        # Calculate Price_Adjusted
        df['Price_Adjusted'] = (
            df['BasePrice'] * df['Quantity'] -
            (df['Quantity'] * df['BasePrice'] * df['Tax_Rate']) -
            (df['Quantity'] * df['Discount'])
        ).astype(int)

        return df

    def calculate_total_sales(self, df):
        '''
        Excel Formula: =SUM(Sales!D2:D6)
        Dependencies: Sales.Price_Adjusted
        '''
        df['Total_Sales'] = df['Price_Adjusted'].sum()
        return df

    def calculate_global_margins_value(self, df):
        '''
        Excel Formula: =A2-(A2*Global_Margins!B2)
        Dependencies: Global_Margins.Value, Financials.Total_Sales
        '''
        df['Global_Margins_Value'] = df['Total_Sales'] - (df['Total_Sales'] * self.global_margins_df['Value'].iloc[0])
        return df

    def calculate_profit_loss(self, df):
        '''
        Excel Formula: =IF(B2>Global_Margins!A2, "Profit", "Loss")
        Dependencies: Global_Margins.Global_Margin, Financials.Global_Margins_Value
        '''
        df['Profit_Loss'] = np.where(df['Global_Margins_Value'] > self.global_margins_df['Global_Margin'].iloc[0], "Profit", "Loss")
        return df

    def transform(self, all_sheets_dict):
        sales_df = all_sheets_dict['Sales']

        # Step 1: Calculate Price Adjusted
        sales_df = self.calculate_price_adjusted(sales_df)

        # Step 2: Calculate Total Sales
        sales_df = self.calculate_total_sales(sales_df)

        # Step 3: Calculate Global Margins Value
        financials_df = self.calculate_global_margins_value(sales_df)

        # Step 4: Calculate Profit or Loss
        financials_df = self.calculate_profit_loss(financials_df)

        return financials_df

# Example usage:
# products_df = pd.DataFrame(...)
# tax_df = pd.DataFrame(...)
# discounts_df = pd.DataFrame(...)
# global_margins_df = pd.DataFrame(...)
# all_sheets_dict = {'Sales': sales_df}

# model = ExcelModel(products_df, tax_df, discounts_df, global_margins_df)
# result_df = model.transform(all_sheets_dict)