# Tariff_Data
This repo contains the code that are  used to collect the data  for Tariff detail .

WORKFLOW 
  Step 1 -->Get Tariff PDF from onlin (By research ) .
  Step 2 -->Run  markit.py to generate .MD file .
  Step 3 --> Load the .MD to the LLM and Give the prompt to extract the data as plain text(bcz LLM strugels to generate an xlsx or ccsv file ) .
  Step 4 --_>XLSX_generator.py ( Make a python code that puts the data into an XLSX file  ) .
