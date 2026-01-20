# test_excel_calculator.py

import pytest
import pandas as pd
import numpy as np
from model import ExcelCalculator

@pytest.fixture
def sample_data():
    # Create sample dataframes with float types for numeric columns
    df = pd.DataFrame({
        'TransactionID': [1.0, 2.0],
        'ProductID': ['P1', 'P2'],
        'Quantity': [10.0, 5.0]
    })

    products_df = pd.DataFrame({
        'ProductID': ['P1', 'P2'],
        'BasePrice': [100.0, 200.0]
    })

    tax_df = pd.DataFrame({
        'Products': ['P1', 'P2'],
        'Tax_Rate': [0.1, 0.2]
    })

    discounts_df = pd.DataFrame({
        'Product': ['P1', 'P2'],
        'Discount': [5.0, 10.0]
    })

    sales_df = pd.DataFrame({
        'Price_Adjusted': [850.0, 800.0]  # Example values
    })

    global_margins_df = pd.DataFrame({
        'Global_Margin': [5000.0],
        'Value': [0.09]
    })

    all_data = {
        'Products': products_df,
        'Tax': tax_df,
        'Discounts': discounts_df,
        'Sales': sales_df,
        'Global_Margins': global_margins_df
    }

    return df, all_data

def test_price_adjusted(sample_data):
    df, all_data = sample_data
    calculator = ExcelCalculator(df, all_data)
    expected = pd.Series([850.0, 800.0])  # Example expected values
    actual = calculator.price_adjusted
    assert actual.equals(expected), f"Failed! Expected: {expected.tolist()}, but got: {actual.tolist()}"

def test_total_sales(sample_data):
    df, all_data = sample_data
    calculator = ExcelCalculator(df, all_data)
    expected = 1650.0  # Sum of the example Price_Adjusted values
    actual = calculator.total_sales
    assert actual == expected, f"Failed! Expected: {expected}, but got: {actual}"

def test_global_margins(sample_data):
    df, all_data = sample_data
    calculator = ExcelCalculator(df, all_data)
    total_sales = 1650.0
    global_margin_value = 0.09
    expected = total_sales - (total_sales * global_margin_value)
    actual = calculator.global_margins
    assert actual == expected, f"Failed! Expected: {expected}, but got: {actual}"

def test_profit_loss(sample_data):
    df, all_data = sample_data
    calculator = ExcelCalculator(df, all_data)
    global_margin_threshold = 5000.0
    global_margins = calculator.global_margins
    expected = "Loss" if global_margins <= global_margin_threshold else "Profit"
    actual = calculator.profit_loss
    assert actual == expected, f"Failed! Expected: {expected}, but got: {actual}"