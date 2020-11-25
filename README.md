Various scripts to pull PA candidate/election info down from the web and save them as well-formatted Excel spreadsheets.

1. `pavoterservices.py` pulls demographic and contact info for all current candidates listed on https://www.pavoterservices.pa.gov/ElectionInfo/ElectionInfo.aspx

2. `montco.py` pulls down precinct-level election results (vote totals only) for the current election at https://electionresults-montcopa.hub.arcgis.com/

3. `lancaster.py` pulls down election summary results from any of the elections listed at "http://vr.co.lancaster.pa.us/ElectionReturns/Election_Returns.html"

Install requirements.

`pip install -r requirements.txt`

Run.

`python {script name}.py`


Results will be saved in the current directory.