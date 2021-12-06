# Analyze Evergreen YT Videos Macro
 Using Excel, determines amount of videos of different age ranges that are above an arbitrary performance level. Allows you to get an idea of how well your videos' views hold up over time. 

### Output:
<img width="662" alt="Screenshot 2" src="https://user-images.githubusercontent.com/12518330/144888313-6a532c36-2cbe-4cdf-a51e-bac8d0ec79cb.png">



### Setup - Instructions
1. Make sure the Developer Tab in Excel is enabled: [Instructions Here](https://support.microsoft.com/en-us/topic/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45)
2. In the Excel Developer tab, click the "Visual Basic" box/button.
3. At the top left, click File > Import File > Select and open the .bas file from this project
4. Close the Visual Basic window, now in the Developer Tab click "Macros" > Select the one that says "Analyze_Evergreen_Videos" and Run

### Acquiring Necessary Data - Instructions:
1. From YouTube Studio Analytics go to 'Advanced Mode'
2. Select a time range of 28 Days or the previous month. 
3. Ensure the 'Views' metric/column is enabled (should be by default)
4. At the top left click the Download icon
5. Select 'Comma Separated Values' - Extract the zip file
6. Open the "Table Data" file and run the macro

### Using the Macro
* The macro will calculate how many videos receive at least a certain number of views, and within certain age timeframes. 
	* For example, you can see what percentage of your YouTube videos that are 6-12 Months old get over 1000 Views/Month and 3000 Views/Month, as well as for 13-24 Months and 25-26 Months.  **(See output screenshot above)**
* When running the macro, it will ask you for these two view counts, which can be anything you like.

#### Limitations:
* The YouTube analytics dashboard will not export more than 500 videos, and it exports using the currently displayed metrics, which sort by Views high-low by default. Therefore, if you have more than 500 videos, *and* you have any videos newer than 36 months that are *not* in the top 500 videos in terms of views, they will not be counted in this metric.
