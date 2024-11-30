# Optimizing Order Fulfillment with Excel Dashboards and Automating KPIs via Office Scripts

## Table of Contents

- [Business Problem](#business-problem)
- [Rationale for the Project](#rationale-for-the-project)
- [Aim of the Project](#aim-of-the-project)
- [Data Description](#data-description)
- [Tech Stack](#tech-stack)
- [Project Scope](#project-scope)
- [Insights and Recommendations](#insights-and-recommendations)
- [Implementation Pathway](#implementation-pathway)

## Business Problem
Streamline Logistics Solutions encounter several pressing challenges within their order fulfillment process, namely:
- Mounting Order Backlogs: Their current routing and resource allocation processes have led to a growing backlog of orders. This situation is compromising delivery timelines and diminishing customer satisfaction.
- Visibility Gap: Customers frequently lack access to real-time updates on their order progress, resulting in communication gaps and increasing dissatisfaction.
- Customer Frustration: The frequency of customer complaints regarding delayed deliveries and inadequate communication channels is rising, tarnishing the company’s reputation for reliability.
- Escalating Costs: Operational expenses are increasing due to overtime payments and the need for expedited shipping to address order backlogs.

## Rationale for the Project
Order Fulfillment is the process of receiving, processing, and delivering customer orders. It involves activities such as inventory management, order processing, picking and packing products, and shipping them to customers. In the context of order fulfillment, backlogs refer to a situation where there is a delay or accumulation of unprocessed orders. 

Backlogs can result from various factors, including high demand, operational inefficiencies, or unforeseen disruptions, and they can negatively impact customer satisfaction, as customers may experience delays in receiving their orders. Eliminating backlogs is crucial to ensuring timely and reliable order fulfillment.

Order fulfillment is the backbone of the company's  operations in the Logistics and Supply Chain industry, where efficiency is not merely a goal but a necessity. Hence, some of the reasons this project is vital for Streamline Logistics Solutions are:

- Customer Satisfaction: Enhancing their order fulfillment processes directly translates into heightened customer satisfaction, thereby nurturing loyalty and long-term relationships.
- Operational Efficiency: Improved efficiency leads to cost savings and heightened profitability, bolstering their competitive position within the industry.
- Data-Driven Insights: Harnessing data-driven insights empowers the business to optimize resource allocation and routing, ensuring timely deliveries and improved resource management.
- Reputation Management: Addressing these operational challenges is paramount to preserving Streamline Logistics Solutions' sterling reputation for delivering excellence consistently.

## Aim of the Project
This project's primary objectives are to develop an Excel interactive dashboard that provides unparalleled visibility into our order fulfillment processes. Also, in order to adequately track and monitor some important metrics important to the business, an automated ad-hoc report will be generated to continously monitor these metrics. Through this, we aim to:
- Efficiently allocate delivery resources based on order volume and location.
- Monitor order progress and proactively identify potential delays.
- Enhance customer communication with timely delivery status updates.
- Reduce order backlogs and operational costs.
- Elevate overall customer satisfaction and safeguard our reputation as an industry leader.

## Data Description
This case study contains 4 datasets and they are as follows;
- Customer Data:
  - Customer_ID: A unique identifier for each customer.
  - Age (years): The age of the customer in years.
  - Gender: The gender of the customer (e.g., Male, Female, Other).
  - Income ($): The annual income of the customer in dollars.
  - Geographic Location: The customer's geographic location (e.g., city, state).

- Sales Data:
  - Transaction_ID: A unique identifier for each sales transaction.
  - Customer_ID: The identifier linking each sale to a customer.
  - Product SKU: Unique identifier for each product.
  - Quantity Sold (units): The number of units sold for each product.
  - Timestamp: The date and time of each sales transaction.
 
- Inventory Data:
  - Product SKU: Unique identifier for each product, linking to sales data.
  - Current Inventory Level (units): The number of units of each product currently in inventory.
  - Stockout (days): The number of days a product has been out of stock.
  - Replenishment Lead time (days): The number of days it takes to replenish inventory.
  - Storage Location: The location where the product is stored.
  - Shelf Life (days): The number of days a product can be stored before expiration.
 
- Production Data:
  - Product SKU: Unique identifier for each product, linking to sales and inventory data.
  - Production Schedule_ID: A unique identifier for each production schedule.
  - Lead Time (days): The time required for manufacturing and distribution.
  - Production Capacities (units per hour): The number of units that can be produced per hour.
  - Resource Allocation: Information about the allocation of resources for production.

## Tech Stack
Tool– Microsoft Excel
- Utilized for creating the interactive dashboard, data visualization, and reporting.
- Data Processing Tools: Leveraging Excel's data manipulation and analysis functions.
- Visualization Tools: Employing Excel's charts, graphs, and pivot tables for order and delivery data visualization.

## Project Scope
- Exploratory Data Analysis: Explore the data to understand its characteristics and discover patterns. 
- Data Analysis: This includes the Customer segmentation, profiling, production planning and performance monitoring.
- Visualization and Reporting: Create interactive dashboards in Excel to visualize insights.  
- Recommendation: Integrate the data-driven supply chain planning process into the company's existing systems and processes.

## Insights and Recommendations
![](Overview.png)


You can view and interact with the dashboard [here](https://hullacuk-my.sharepoint.com/personal/m_o_lawal-2021_hull_ac_uk/_layouts/15/Doc.aspx?sourcedoc={82d09735-b334-48bd-ba3d-d88c401bae9f}&action=embedview&AllowTyping=True&ActiveCell='Dashboard'!A1&wdHideGridlines=True&wdHideHeaders=True&wdInConfigurator=True&wdInConfigurator=True)

## General Insights

1. Diverse Customer Base: The even distribution across different ages and the presence of various gender groups indicate a wide-ranging customer base.

2. Income and Age Group Correlation: The majority of customers being adults with High income suggests priority to SKUs for this customer sect.

3. Geographic Variability: Different preferences in different cities indicate the need for region-specific strategies.

4. Product Preferences by Segment: Each customer segment (Type 1 to Type 6) shows distinct preferences for certain SKUs, highlighting the importance of tailored product offerings.

5. High Value SKUs: The identification of top SKU bought by customers suggests a focused demand for specific products.

6. Sales Variability and Seasonality: Sales varies for different month when in different locations by different Customer Types.


## Recommendations

1. Product Strategy:
   - Diversification: Develop products catering to the diverse needs of different age and income groups.
   - Tailoring for High-Value SKUs: Focus production on top SKU from the customer segment, ensuring they are well-stocked and readily available.

2. Inventory Management:
   - Dynamic Stocking: Adjust inventory levels based on sales trends of high-value SKUs for each customer type.
   - Regional Focus: Tailor inventory to regional preferences, ensuring that popular products in each city are adequately stocked based customer preferences.

3. Marketing and Sales:
   - Segment-Specific Campaigns: Develop targeted marketing campaigns for each customer segment (Type 1 to Type 6), focusing on their preferred SKUs.
   - Promotional Timing: Align marketing campaigns with peak sales periods to capitalise on increased demand.

4. Production Planning:
   - Demand-Driven Production: Adjust production schedules based on the sales trends of the top SKUs which are the specific demands of each customer segment.
   - Efficiency in Low-Demand Products: For the lower-value SKUs, which are the least sort after SKUs, adopt a lean production approach to minimise excess inventory.

5. Regional Strategies:
   - Localized Offerings: Customize product offerings and stock levels based on the preferences in Cities A, B, and C.
   - Distribution Optimization: Optimize distribution and logistics to cater to the unique demands of each geographic location.

6. Data-Driven Approach:
   - Continuous Monitoring: Regularly analyze sales data to adapt strategies in real-time.
   - Feedback Loops: Use customer feedback to refine product offerings and marketing strategies.

7. Customer Engagement:
   - Community Building: Engage with different customer groups to build brand loyalty and gain deeper insights into their preferences.
   - Responsive Customer Service: Provide excellent customer service, especially to the dominant adult low-income segment, to foster long-term relationships.

## Implementation Pathway

- Short-Term: Focus on inventory adjustments and marketing campaigns targeting key segments and high-value SKUs.
- Medium-Term: Develop and refine products based on customer segment feedback and sales data analysis.
- Long-Term: Invest in technology and processes that enhance understanding of customer preferences and market dynamics, enabling agile and responsive supply chain management.

By adopting these recommendations, SmartHome Solutions Inc. can better align its products and services with the diverse needs of its customer base, optimising both customer satisfaction and operational efficiency.
