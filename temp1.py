import sys, os
with open("/Users/divakar/Desktop/MicrosoftFileReader/daily_excel_processor.log", "a") as f:
    f.write("Python: {}\n".format(sys.executable))
    f.write("Sys Path: {}\n".format(sys.path))
    f.write("Env PATH: {}\n".format(os.environ.get("PATH")) + "\n")
