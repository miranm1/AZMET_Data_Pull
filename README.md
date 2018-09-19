This program uses Microsoft Excel VBA to pull raw data from the AZMET weather station for certain locations. It is pulling data for precipitation, reference et value, and reference et "Penman Monteith". It is separated into three different spreadsheets, but it is all working from one macro module. This was done by creating three different sub procedures and sub functions with different names for the corresponding data being obtained. The program works getting the start year, end year, and the station number given from the website.  It will then grab the raw comma delimited data from the built URL. This data is copied into a temp spreadsheet to and separated comma delimited format. It will then go down the column value by value  and place it into the main spreadsheet. When the column ends it goes to the next weather station column and repeats until there is no more data in the “stationNum:” row on the main spreadsheet. 

# How to use:


To start the program just input the start and end year on C1 and C2. When inputting the year for cell C1 and C2 use the last two digits of the year like the picture above.  This is because the values are used for the URL builder in the macro. Now just press the button to start the program. 



More stations can be added to the program by just putting the assigned value for the weather station into the next empty column in row 6. If the number starts with a zero make sure to the the cell into a text format.  This will ensure the URL is written correctly.  To obtain this number just go to the website https://cals.arizona.edu/azmet/az-data.htm and select the station you want to add to the program. 


After selecting the station you will be able to see the number of the station in the URL. The number in case of the picture about is 07. 



When you press the button the spreadsheet should populate shortly like this picture above. 

