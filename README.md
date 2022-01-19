# DataModelDoc
Repository for Model Documentation Automation


## RPD setup and data download
To use this script you will first need to get all data from the repository

After opening your BI Administration Tool, open the RPD using the Cloud option (BI administration tool -> Open -> Cloud) and log in:

![DataModelDocScreenshot_](https://user-images.githubusercontent.com/26796318/150189173-4453e422-9e53-49d0-88ea-612f200be1b6.png)


After opened, you can download the CSV and the XML files using the utilities tool

![DataModelDocScreenshot2](https://user-images.githubusercontent.com/26796318/150190336-f8fe4016-f0e6-4601-b2b8-e78d814e6dba.png)

CSV file:
  Tools -> Utilities... -> Repository Documentation
  
XML file
  Tools -> Utilities... -> Generate Logical Column Type Document
  
![DataModelDocScreenshot3](https://user-images.githubusercontent.com/26796318/150190462-9c0f540d-9df3-40c3-8f55-d89189ab44f4.png)


## Data Model Documentation tool

Running the Makefile script, you should get all requiriments installed as well as the python version needed.
It will also run the GUI.py script to get this window

![image](https://user-images.githubusercontent.com/26796318/150193170-1e7e5e32-a47f-405b-a345-bd324a8ed2b8.png)

Here you can add those files that you got from the RPD, the CSV and the XML, through the 'Browse' buttons.
Add also the name that you want for the generated file and your name.

After this setup, the new documentation workbook will be generated in the same folder where the CSV file is and opened up by clicking 'RUN'
