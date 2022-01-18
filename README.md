# HKN Beta Chapter Recruitment Filter
This is a lil script that outputs the list of names of eligible candidates from the pool of current ECE students given our requirements. 


## Setup
This script relies on at least one or more additional packages other than the standard ones.
To install them use the `requirements.txt` file. 
You can run for example
```
pip install -r requirements.txt
```


## Instructions 
1. Run Cognos 'record copy - PUID entry' report for **all ECE students**. (Note that these must be processed in batches, as Cognos has a limit of 1000 entries at once per report.)
2. When Cognos reports are done in batches, put all the excel files (make sure has .xlsx extension) in same folder as the script.
3. Run the python script. Make sure the script and the excel files are in the same working directory.
4. It should output 3 csv files called `seniors.csv`, `juniors.csv`, and `sophomores.csv` with the names and PUIDs of the eligible candidates.
5. Send these 3 files to the HKN faculty or staff advisor or whoever the point of contact is.


## Troubleshooting
If there is any difficulty in running the script, submit an issue on the Github page or contact your hkn beta chapter representative with the stack trace and error message and a description of what you were doing.
