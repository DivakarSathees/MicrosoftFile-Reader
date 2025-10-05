#!/bin/bash
export PATH=/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin
export PYTHONPATH=$HOME/Library/Python/3.9/lib/python/site-packages

# optional: set other environment variables needed by your script
export ACCESS_TOKEN=""
# run the Python script
/usr/local/bin/python3 "/Users/divakar/Desktop/MicrosoftFileReader/dateFetch.py"
