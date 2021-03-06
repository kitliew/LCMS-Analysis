# LCMS-Analysis
Generate summary results and graphs from Xcalibur-QuanBrowser Excel Report

Disclaimer: This is a personalized script with limited knowledge on automated analysis for LC-MS.


Xcalibur™ Software - Thermo Fisher Scientific is a powerful software. It is used in our lab to control and process Liquid chromatography–mass spectrometry (LC–MS) data.

LC-MS typical work flow:
1. Run LCMS samples
2. Modify and save the following as new sequence for analysis (generate analysis.sld):
      a. Sample ID (grouping)
      b. Proc Meth (Processing setup)
      c. Sample wt (optional)
3. Batch reprocess
4. QualBrowser & QuanBrowser for data analysis
5. Generate Report.xls (short or long excel report generate from QuanBrowser; Long Report suppose to generate "Sample wt" column, but error/blanks occur very often).
6. Go through each sheet in ShortReport.xls to 
      a. normalized data
      b. calculate mean and standard deviation of each group
7. Create Box and Whisker plot with individual data to observe distribution of data.
8. Create Bar Graph with Standard deviation as Error bar.  


Step 6 - 8 is repetitive and time consuming for large dataset.


Requirement:
1. Report file in .xlsx format
2. Weight file in .xlsx format, containing "Filename" column and "Sample wt" column (position/loc sensitive).
Templates/reference attached in Example folder.
 
## How To Use?
1. Run AnalysisGenerator.py (Python required) **OR** "AnalysisGenerator Beta" > dist > AnalysisGenerator > AnalysisGenerator.exe
2. Input Internal Standard concentration
3. Select weightfile.xlsx
4. Select reportfile.xlsx
5. A Result folder will be generated in the same directory

![alt text](https://github.com/kitliew/LCMS-Analysis/blob/main/Example/Results/Summary_graph_example.PNG)

![alt text](https://github.com/kitliew/LCMS-Analysis/blob/main/Example/Results/VitaminC.png)


Feel free to use and comment.

