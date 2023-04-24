# Holdings Record Inpsector
Pulls specified data from holdings records in Alma, using the Bibs API

There are three separate versions of the script here:

## Holdings Inspector_sharable.py
Basic script to retrieve specified holdings records (by MMS ID and Holdings ID). Will retreive the following from the 852 field:
- First and second indicators
- $c (location)
- $h (Classification part) (Also looks for unusual spacing or formatting)
- Count of $i (Item part) subfields
- $k (Call number prefix)
- $l (Shelving form of title)

## Holdings Record Inspector_sharable.ipynb
Jupyter Notebook version of same script described above. (Sometimes I find it easier to work with a Jupyter Notebook when looking to check output of new/additional fields, etc.)

## Holdings inspector with selections_sharable.py
More advanced version that prompts user for which of the parameters should be run. The check for each is a separate function, each of which runs if the condition to run it is true.

# You will need:
- Bibs and Inventory API key for your Alma environment (Read-only)
- Excel file containing MMS IDs and Holdings IDs for holdings records you would like data from (I enclose my ID numbers in ""s to prevent any issues with the IDs being misinterpreted as numbers and the script will remove them after importing them as strings)
