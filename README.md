# Optimizing Order Fulfillment with Excel Dashboards and Automating KPIs via Office Scripts

## Table of Contents

- [Business Problem](#business-problem)
- [Rationale for the Project](#rationale-for-the-project)
- [Aim of the Project](#aim-of-the-project)
- [Data Description](#data-description)
- [Tech Stack](#tech-stack)
- [Project Scope](#project-scope)
- [Insights and Recommendations](#insights-and-recommendations)

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
The data for this project comprises of the following;

- Order ID: A unique identifier for each customer order.
- Delivery Address: The address to which the order is to be delivered.
- Order Timestamp: The date and time when the order was placed (e.g., "2023-09-01 08:00").
- Order Status: The current status of the order (e.g., "In Progress" or "Completed").
- Driver ID: A unique identifier for each driver assigned to deliver orders.
- Vehicle Info: Information about the delivery vehicle used for the order.
- Current Location: The current location of the delivery driver during order delivery.
- Delivery Time: The total time taken for delivery, measured in minutes (e.g., "120 min").
- Delays: Any delays that occurred during the delivery, measured in minutes (e.g., "15 min").
- Customer Feedback: Feedback from the customerregarding the delivery experience (e.g., "Positive" or "Negative").
- Route: The specific route taken by the delivery driver for the order.
- Delivery Zone: The geographic zone or area where the delivery is made.
- Allocation Rules: Rules used to allocate resources for the delivery (e.g., "Standard Rules" or "Expedited Rules").
- Timestamp for Tracking: The date and time of tracking data points (e.g., "2023-09-01 08:15").

## Tech Stack
Tool– Microsoft Excel
- Utilized for creating the interactive dashboard, data visualization, and automated ad-hoc report generation using office scripts.
- Data Processing Tools: Leveraging Excel's data manipulation and analysis functions.
- Visualization Tools: Employing Excel's charts, graphs, and pivot tables for order and delivery data visualization.

## Project Scope
1. Data Preprocessing and Exploratory Analysis
   - Rigorous data formatting and preparation
   - Identify patterns, correlations, and potential data anomalies
   - Address any data quality issues

2. Data Augmentation
   - Enrich the original dataset with additional relevant information
   - Prepare data to comprehensively answer business questions

3. Data Visualization and Dashboard Design
   - Create visual representations of key insights
   - Develop an interactive Excel dashboard with intuitive data visualization components
   - Effectively communicate findings through clear and compelling visuals

4. Automated Reporting
   - Create Office Scripts to automate ad-hoc report generation
   - Develop scripts to calculate and track important business metrics
   - Ensure consistent and efficient metric tracking

5. Interpretation and Recommendations
   - Extract meaningful insights from the analyzed data
   - Develop comprehensive documentation
   - Provide actionable recommendations based on the analysis

## Insights and Recommendations
![](Overview.png)


You can view and interact with the dashboard [here](https://hullacuk-my.sharepoint.com/personal/m_o_lawal-2021_hull_ac_uk/_layouts/15/Doc.aspx?sourcedoc={82d09735-b334-48bd-ba3d-d88c401bae9f}&action=embedview&AllowTyping=True&ActiveCell='Dashboard'!A1&wdHideGridlines=True&wdHideHeaders=True&wdInConfigurator=True&wdInConfigurator=True)

## Insights

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

