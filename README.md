# HKN Beta Chapter Recruitment Filter
This is a lil script that outputs the list of names of eligible candidates from the pool of current ECE students given our requirements. 


## Instructions 
1. Run Cognos 'record copy - PUID entry' report for **all ECE students**. (Note that these must be processed in batches, as Cognos has a limit of 1000 entries at once per report.)
2. When Cognos reports are done in batches, put all the excel files (make sure has .xlsx extension) in same folder as the script.
3. Run the python script. Make sure the script and the excel files are in the same working directory.
4. It should output 3 csv files called `seniors.csv`, `juniors.csv`, and `sophomores.csv` with the names and PUIDs of the eligible candidates.
5. Send these 3 files to the HKN faculty or staff advisor or whoever the point of contact is.
