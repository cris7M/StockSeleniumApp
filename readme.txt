Requirment
    Python3 and pip library
    
Create a virtual Enviroment In linux
    virtualenv <virtualenv_name>

Activate virtual Enviroment
    source virtualenv_name/bin/activate

Installation of python library 
    There is a file named scrapRequirment.run this file with below command
    pip install -r  scrapRequirment.txt
    it will install all the required library to run the script

FILES IN THE FOLDER AND THERE RESPONSIBILITIES:
    1. InputData.xlxs - This excel file has all the name of the company for which we have to scrap data.
    2. ExcelUtils.py - This python file has one method to make all possible combination of the comapnay name taken from InputData.xlxs.
    3. central_scrapping_FINAL.py - This python file will extract all the data from central for the company name 
    which matches from InputData.xlxs and keep that data into excel file, these excel file will being dump to database.
    4. state_scrapping_FINAL.py - This python file will extract the data of all state one by one and store it to excel file, 
    these excel file will being dump to database.

Log Files 
    Log file will being recorded sepratly for central and state with different 
    folders named central_log and state_log respectively.
Error Files
     Log file will being recorded sepratly for central and state with different 
     folders named central_error_log and state_error_log respectively.
Output Files    
    Output files in excel format will being recorded sepratly for central and state with different 
    folders named central_output and state_output respectively.
