import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

conn = pyodbc.connect(
    r"DRIVER={ODBC Driver 18 for SQL Server};"
    r"SERVER=localhost;"
    r"DATABASE=AdventureWorks2019;"
    r"Trusted_Connection=yes;"
    r"Encrypt=no;"
)

query = """
SELECT 
    soh.SalesOrderID,
    soh.OrderDate,
    soh.CustomerID,
    soh.SubTotal,
    soh.TaxAmt,
    soh.Freight,
    st.Name AS Territory
FROM Sales.SalesOrderHeader soh
LEFT JOIN Sales.SalesTerritory st ON soh.TerritoryID = st.TerritoryID
"""
df = pd.read_sql(query, conn)
conn.close()

print(df.head())
print(df.shape)
print(df.columns)

print("Before drop ", df.isna().sum())
df.dropna(inplace=True)
print("Missing values after drop:", df.isna().sum())
print("Shape after dropping NA:", df.shape)

print("Number of duplicate rows:", df.duplicated().sum())
df.drop_duplicates(inplace=True)
print("Number of duplicate rows after drop:", df.duplicated().sum())
print("Shape after dropping duplicates:", df.shape)

numeric_cols = df.select_dtypes(include=[float, int]).columns
print("Numeric columns:", numeric_cols)
print("Shape before outlier removal:", df.shape)
plt.boxplot(df[numeric_cols])
plt.title('Boxplot of Numeric Columns')
plt.show()

plt.figure(figsize=(10, 8))
corr_matrix = df[numeric_cols].corr()
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', fmt=".2f")
plt.title('Correlation Heatmap of Numeric Features')
plt.show()

for col in numeric_cols:
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    df = df[(df[col] >= lower_bound) & (df[col] <= upper_bound)]

print("Shape after outlier removal:", df.shape)
plt.boxplot(df[numeric_cols])
plt.title('Boxplot of Numeric Columns After Outlier Removal')
plt.show()

print("Data after outlier removal:")
print(df.head())

df['TotalSales'] = df['SubTotal'] + df['TaxAmt'] + df['Freight']
df['Year'] = pd.to_datetime(df['OrderDate']).dt.year

plt.hist(df['TotalSales'], bins=40, color='green', edgecolor='black')
plt.title("Total Sales Distribution")
plt.xlabel("Total Sales")
plt.ylabel("Count")
plt.tight_layout()
plt.show()

df.groupby('Year')['TotalSales'].sum().plot(kind='bar', color='skyblue')
plt.title("Total Sales by Year")
plt.ylabel("Revenue")
plt.xticks(rotation=0)
plt.tight_layout()
plt.show()

df.groupby('Territory')['TotalSales'].sum().sort_values().plot(kind='barh', color='orange')
plt.title("Sales by Territory")
plt.xlabel("Revenue")
plt.tight_layout()
plt.show()

sns.heatmap(df[['SubTotal','TaxAmt','Freight','TotalSales']].corr(), annot=True, cmap='coolwarm')
plt.title("Sales Feature Correlations")
plt.tight_layout()
plt.show()

print("Orders with unusually high or low Total Sales:")
print(df[df['Outlier_TotalSales']][['SalesOrderID', 'TotalSales']].sort_values(by='TotalSales', ascending=False))