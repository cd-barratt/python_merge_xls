import numpy as np
import pandas as pd
import glob
import os
from datetime import datetime

print '\n', "======================= *.xls file merger for python 2.7 =======================", '\n'
print "This script will take any number of *.xls files in a directory and merge them according to their common fields. Ensure that files are in the same working directory as this script, and make sure the outputfile is changed to a relevant name", '\n'

#change outputfile here
outputfile = 'MERGED FILES.xlsx'

#count xls files
filecount = int(0)
for filetype in os.listdir('.'):
    if filetype.endswith('.xls'):
        filecount += 1
print "> There are {} *.xls files in this directory that will be merged.".format(filecount)
print "> Output file will be named: ", outputfile, "- please edit line 12 of this script if you want to change this"
raw_input("> If you wish to go ahead please press <Enter> to continue...")

#start clock
t_begin = datetime.now()


# merge the files
print "> Merging files now..."
all_data = pd.DataFrame()

for f in glob.glob("*.xls"):
	df = pd.read_excel(f, "Detailed survey data") # this is the tab of the sheet you want to merge
	df['filename'] = os.path.basename(f)
	all_data = all_data.append(df,ignore_index=True)

# now save the data frame
writer = pd.ExcelWriter(outputfile)
all_data.to_excel(writer,'MERGED')
writer.save()
print "> Saving to merged file:", (outputfile), "..."

#clock it!
t_finish = datetime.now()
elapsed = t_finish - t_begin

print '\n', '\n', "-----------------------------------------------------------------------------------------------------"
print "> Completed merging!"
print "Total time: {0} (H:M:S)".format(elapsed)
print "-----------------------------------------------------------------------------------------------------", '\n', '\n'
