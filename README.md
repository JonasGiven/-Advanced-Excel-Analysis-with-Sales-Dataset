
## Advanced Excel Analysis with Sales Dataset

### Project Description
This project involves advanced Excel techniques to analyze a sales dataset. The analysis includes creating pivot tables, using array formulas, visualizing data with advanced charts, and performing what-if analysis to forecast decision-making with projected budgets.

### Requirements
- Microsoft Excel
- [Sales dataset](https://witscloud-my.sharepoint.com/:x:/g/personal/2168978_students_wits_ac_za/EUdSzYo6vnVKjp7EA9TMQY0BNWGA4tWtOor8H1VDsvpxKg?e=UTduge)

### Technologies Used
- Microsoft Excel

### Methods

#### 1. Pivot Tables
**Analyzing sales dataset for top-performing products and sales trends**

- **Create Pivot Table:**
  - Select the dataset.
  - Navigate to the **Insert** tab.
  - Choose **PivotTable**.
  - Select a range for the PivotTable and place it in a new worksheet.

- **Summarize Total Sales by Product:**
  - Drag the **Product** field to the Rows area.
  - Drag the **Sales** field to the Values area to display total sales per product.

- **Group Sales Data by Month and Year:**
  - Right-click a date in the pivot table.
  - Select **Group**, then choose **Month** and **Year** to group the sales data accordingly.

- **Use Slicers for Dynamic Filtering:**
  - Go to the **PivotTable Analyze** tab.
  - Click on **Insert Slicer** and choose **Region**.
  - Use the slicer to filter the data by selecting different regions.

- **Add Calculated Field for Profit Margins:**
  - In the **PivotTable Analyze** tab, select **Fields, Items, & Sets**, then **Calculated Field**.
  - Name it "Profit Margin" and set its formula to `=Profit / Sales * 100`.

#### 2. Array Formulas
**Conducting detailed customer feedback score analysis**

- **Average Feedback Score by Product Category:**
  - Select a cell.
  - Enter `=AVERAGE(IF((category_range="Category Name"), feedback_scores_range))`.
  - Press `Ctrl+Shift+Enter`.

- **Max, Min, and Average Scores for Each Product:**
  - Select multiple cells for the results.
  - Enter `={MAX(range), MIN(range), AVERAGE(range)}` for the selected product.
  - Press `Ctrl+Shift+Enter`.

- **Count Each Score Across All Products:**
  - In a new cell, use `=COUNTIF(scores_range, "=Score")` within an array formula.
  - Press `Ctrl+Shift+Enter`.

#### 3. Advanced Charts (Combo Charts, Tree Maps)
**Visualizing the company's market share data**

- **Create a Combo Chart for Sales Volume vs. Profit Margin:**
  - Select monthly sales and profit margin data.
  - Navigate to the **Insert** tab.
  - Click **Combo Chart** and choose a suitable combo chart type.
  - Customize the chart by setting the sales volume on the primary axis and profit margin on the secondary axis.

- **Use a Tree Map for Product Market Share:**
  - Organize data with product categories and sales.
  - Select the data.
  - Insert a Tree Map chart from the **Insert** tab.
  - Customize the tree map with labels for product names and sales values.

- **Customize Charts:**
  - Add legends by selecting the chart and using the Chart Elements button.
  - Add titles and data labels for clarity.

#### 4. What-if Analysis
**Forecasting decision-making with projected budget**

- **Goal Seek for Target Profit:**
  - Navigate to the **Data** tab.
  - Select **What-If Analysis**.
  - Click **Goal Seek**.
  - Set the target value (profit) and change the cell (sales volume) to determine the required sales volume.

- **Two-Variable Data Table for Sales Volume and Unit Cost Impact:**
  - Set up a table with different sales volumes and unit costs as row and column headers.
  - Use the Data Table function under **What-If Analysis** to fill the table.

- **Scenario Manager for Budget Scenarios:**
  - Under the **Data** tab, choose **What-If Analysis**, then **Scenario Manager**.
  - Create scenarios named "Best-Case", "Worst-Case", and "Most-Likely" with varying sales and cost values.
  - Compare the outcomes by switching between scenarios.

### Results

See below a snapshot of the final outcome.
![Screenshot 2024-05-19 at 06 44 04 2024-05-19 04_44_52](https://github.com/JonasGiven/-Advanced-Excel-Analysis-with-Sales-Dataset/assets/169194581/a3358ec9-53d6-45bd-93af-27f0352dd2d0)
To access the excel document [click here](https://witscloud-my.sharepoint.com/:x:/g/personal/2168978_students_wits_ac_za/EWrsXQKuuSZIkJOVXOGKg5cBIIsEwLcrG3OiQgbrOUstJA?e=0vhSq0).<br/>
In order to open the file your device must atleast have Microsoft Excel or any Excel file reader software or copy the link and paste it on your browser (Google chrome, Microsoft Edge, Safari, Opera etc)
---

