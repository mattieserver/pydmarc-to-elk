#  pyDMARC-to-ELK

pyDMARC-to-ELK is a python tool that connects to a mailbox and send the data to elastic.

## Setup

This tool assumes that only dmarc reports are in the mailbox.
You need to first create the config file.
You can do this with:

python3 writedefaultconf.py

After that you can edit the Settings/config.ini file and enter the correct credentials.


Once the config file is set you can run the tool with:

python3 pyStart.py 

