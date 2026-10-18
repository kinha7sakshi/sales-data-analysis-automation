import pandas as pd

df = pd.read_excel("sales.xlsx", engine="openpyxl")

# Clean column names
df.columns = df.columns.str.strip()

# Clean product names
df["Product"] = df["Product"].fillna("").astype(str).str.strip()

# Clean customer names
df["Customer"] = df["Customer"].fillna("").astype(str).str.strip()
# Calculate sales
df["Sales"] = df["Quantity"] * df["Price"]

# Calculate total sales
total_sales = df["Sales"].sum()

# Sales by product
product_sales = df.groupby("Product")["Sales"].sum()

# Find best-selling product
best_product = product_sales.idxmax()
best_product_sales = product_sales.max()

customer_sales = df.groupby("Customer")["Sales"].sum()

top_customer = customer_sales.idxmax()
top_customer_sales = customer_sales.max()
df.to_excel("automated_sales_report.xlsx", index=False)
# Create summary information
summary = {
    "Metric": [
        "Total Sales",
        "Best-Selling Product",
        "Best Product Sales",
        "Top Customer",
        "Top Customer Sales"
    ],
    "Value": [
        total_sales,
        best_product,
        best_product_sales,
        top_customer,
        top_customer_sales
    ]
}

summary_df = pd.DataFrame(summary)

# Add summary to the Excel report
with pd.ExcelWriter(
    "automated_sales_report.xlsx",
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:
    summary_df.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )
print(df)
print("Total Sales:", total_sales)
print("Sales by Product:")
print(product_sales)

print("Best-Selling Product:", best_product)
print("Best Product Sales:", best_product_sales)
print("Sales by Customer:")
print(customer_sales)

print("Top Customer:", top_customer)
print("Top Customer Sales:", top_customer_sales)
