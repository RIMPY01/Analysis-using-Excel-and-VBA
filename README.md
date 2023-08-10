# Analysis-using-Excel-and-VBA
The project aims to deliver an efficient and visually appealing dashboard that empowers users to perform in-depth analysis of Netflix's quarterly revenue and subscription data, providing valuable insights for strategic decision-making.

Created a customized Excel tool for comprehensive analysis of the quarterly revenue and subscription data of Netflix Production Company.

Developed a macro to format the Netflix subscription data and created Buttons for Key Metrics: Integrated three buttons to generate separate columns for total revenue, total subscribers, and the percentage quarterly change. These features will provide quick insights and be highly user-friendly.

Sub-Procedures Implementation for Streamline Data analysis:

a. countryWiseYearlyData(countryRow, what): This procedure calculated total yearly subscriber count and revenue per country, offering the flexibility to choose between revenue and subscriber analysis.

b. getQuaterlyDataForGlobe(quarterColumn, what): this procedure is designed to retrieve quarterly data for all countries combined. The parameter "what" will specify revenue or subscriber data, facilitating efficient summation of required cells.

Percentage Change Analysis:
The sub-procedure "Quarter Change Country-wise" Shows percentage changes in quarterly revenue and subscription per country. Based on user input, this procedure will calculate the percentage increase or decrease and present the visualization over quarters appropriately.

Highest and Lowest Metrics:
Created four buttons to identify the countries with the highest and lowest subscriptions as well as the highest and lowest revenue. These buttons will provide quick access to critical insights.

Data Visualization with Charts:
Utilize various charts to present the analysis:
a. Custom World Heat Map: displays yearly subscriptions across the globe.
b. Custom World Heat Map: displays yearly revenue per country
c. Revenue Bar Chart: depicts the greatest and lowest Quarterly revenue as per country.
d. Subscription Change line Chart: depicts the greatest and lowest Quarter subscription changes.
