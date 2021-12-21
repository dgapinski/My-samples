# My-samples
These are samples I’ve made for making my admin life easier:

LARD - LAPS Admin Remote Desktop 
facilitates LAPS RDP connections with GUI and credential lookup. Basically, if you have rights to 
look up the password in AD, then the LARD button turns green. If you don't have rights, it turns 
red. If green, you can click it to open up an RDS connection to that endpoint.

JOBMASTER.PS1
this script runs remote scripts in job form, on all computers in a windows domain OU
that you choose. 
First shows a form for you to pick an OU, and then choose a script to run. All 
scripts live in the \Jobs subfolder, and are listed in the form. Then when the 
job runs on the chosen OU, the final computer list is made of computers that
respond to pinging, and the results output back to grid view when they complete.
