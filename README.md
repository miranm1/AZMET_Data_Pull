# AZMET_Data_Pull
This program uses Microsoft Excel VBA to pull raw data from the AZMET weather station for certain locations. It is pulling data for precipitation, reference et value, and reference et "Penman Monteith". It is separated into three different spreadsheets, but it is all working from one macro module. This was done by creating three different sub procedures and sub functions with different names for the corresponding data being obtained. The program works getting the start year, end year, and the station number given from the website.  It will then grab the raw comma delimited data from the built URL. This data is copied into a temp spreadsheet to and separated comma delimited format. It will then go down the column value by value  and place it into the main spreadsheet. When the column ends it goes to the next weather station column and repeats until there is no more data in the “stationNum:” row on the main spreadsheet. 

## How to use:
1. To use this macro application in Excel just input the years you would like to take data from in start year and end year. Only use the last two digits in the year you are gathering from.
![alt text](https://github.com/miranm1/AZMET_Web_Scrapper/blob/master/azmet1.PNG)

2. The Excel spread sheet will soon populate with precipitation, reference et, or reference et "Penman Monteith" depending on the tab worksheet you are working in.  The units of measurements in the program are in inches(hundredths).
![alt text](https://github.com/miranm1/AZMET_Web_Scrapper/blob/master/azsniplast.PNG)

## Adding more stations:
To add additional stations so the macro application you have to find the station number on the AZMET website and simply input it to the next avaible column. 
![alt text](https://github.com/miranm1/AZMET_Web_Scrapper/blob/master/azmet2.PNG)
