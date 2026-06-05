PROJECT TITLE
Product Dataset Excel Workbook
Overview
This Excel workbook contains a product dataset created for data analysis, reporting, inventory management, and business intelligence purposes.
Workbook Contents
The workbook includes a worksheet named "orderDB" containing product-related information.
Columns Included
- Product_ID: Unique identifier for each product
- Product_Name: Name of the product
- Brand Name: Manufacturer or brand name
- Quantity: Available inventory quantity
- Category: Product category
- Price: Selling price of the product
Data Format
- Product_ID values are unique.
- Prices and costs are stored as numeric values.
- product name as character.
- Brand name  as character
- Category as character
  Purpose
  This dataset can be used for:
- Inventory tracking
- Product performance analysis
- Sales reporting
- Calculate Total Price of All Products
The total of all products is calculated using the SUM function:
- COLUMN D:price($)
- Column B: product name
EXCEL FORMULA:SUM(D2:D35)
- How many product exist in dataset
- Column B contains the Product Name.
- COUNT B counts all non-empty product records.
- Excel formula count(B2:B35)
- Calculate the Averge price of product
- Column D contains the prices.
- AVERAGE calculates the mean value of all prices.
- EXCEL FORMULA AVERAGE(D2:D35)
- Objective
   The purpose of this workbook is to identify the product with the lowest price and calculate the minimum price in the dataset.
- Minimum Price Calculation
  The minimum product price is calculated using the Excel `MIN` function 
    EXCEL FORMULA:MIN(D2:D35)
Objective
   The purpose of this workbook is to identify the product with the highest price and  calculate the MAXMINUM  price in the dataset.
- MAXMINUM  Price Calculation
  The maxminum product price is calculated using the Excel `MAX` function 
    EXCEL FORMULA:MAX(D2:D35)
 Excel Workbook: Logical Function Using IF Condition
-  Overview
        This Excel workbook contains a product dataset and demonstrates the use of logical functions using the Excel `IF` condition. It is designed to help classify products based on specific conditions such as price, Product
- Worksheet: Product_Dataset
Objective
The purpose of this workbook is to apply logical conditions using the `IF` function to categorize or evaluate product data.
- IF Function Logic formula
- =IF(D2:D35,">=$500","High price")
- =if(D2:D35,"<,$500","Standard price")
- Excel Workbook: SUMIF Conditional Calculation (Electronics Total Price)
  Overview
       This Excel workbook contains a product dataset and demonstrates how to calculate the total price of products in a specific category using           the `SUMIF` function. In this case, we calculate the total price of all products in the **Electronics** category.
Where: Product dataset
Column F= Category
Column D = Price
"Electronics" = condition used for filtering
SUMIF formula
=SUMIF(D2:D35,"Electronics",D2:D100)
 Overview
       This Excel workbook contains a product dataset and demonstrates how to calculate the total price of products in a specific category using           the `countif function. In this case, we calculate the total price of all products in the **>$100** category.
where :product dataset
Column D:price
=COUNTIF(D2:D35,">$100")
- Create 2-Character Product ID Using LEFT Function
Overview
       This Excel workbook contains a product dataset and demonstrates how to create a shortened Product ID using the `LEFT` function in Excel. A          new column is created to extract the first 2 characters from the original Product ID.
Worksheet: Order DB
- Objective
    The purpose of this workbook is to generate a **2-character product ID** from the original Product_ID using the Excel `LEFT` function named as Day.
Where:
- Column A = Product_ID
- 2 = Number of characters to extract from the left side
- LEFT Function Formula
- =LEFT(A2,2)
Character Product ID Using RIGHT Function
Overview
       This Excel workbook contains a product dataset and demonstrates how to create a shortened Product ID using the `RIGHT ` function in Excel.         new column is created to extract the first 2 characters from the original Product ID.
Worksheet: Order DB
- Objective
    The purpose of this workbook is to generate a **2-character product ID** from the original Product_ID using the Excel `RIGHT ` function named as Countery code.
Where:
- Column A = Product_ID
- 2 = Number of characters to extract from the Right side
RIGHT Function Formula
- =RIGHT(A2,2)
Excel Workbook: Extract Month from Product ID Using MID Function
- Overview
    This Excel workbook contains a product dataset and demonstrates how to extract specific characters from the Product ID using the Excel `MID`         A new column is created to extract the **month value (4th to 6th characters)** from the Product ID.
Where:
Column A = Product_ID
4 = Starting position (4th character)
6 = Number of characters to extract (4th to 6th)
MID Function formula
=MID(A2,4,6)








- Dashboard creation
- Business analytics and forecasting
