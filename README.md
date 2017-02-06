Brachytherapy Vienna spreadsheet audit
======================================
This is a tool to scrape treatment data from the Vienna brachytherapy [spreadsheet](https://www.americanbrachytherapy.org/guidelines/gyn_HDR_BT_docu_sheets.xls) to make it more easily auditable.

 ## Setup

 ```
pip install -r requirements
 ```

You'll need a local server of [MongoDB](https://www.mongodb.com/) running.

[Ghost Driver](https://github.com/detro/ghostdriver) is required alongside [Selenium](http://selenium-python.readthedocs.io/) to generate HTML files containing interactive plots of retrieved data.

There is the option to correlate data in the spreadsheet to RTPLAN files on an Oncentra MasterPlan server. If you want to use this functionality, modify `omppackage\server_config.cfg` to your server specifications.

 ## Usage

 Make a copy of the directories containing the spreadsheets in `/static`. Then run

 ```
python main.py
 ```
 to crawl through `/static` and add data into a MongoDB collection.

 To query the database and make some visualisations you can run:

 ```
 python retrieve.py
 ```

 but you'll probably want to dig down into `retrieve.py` and modify it to your requirements. Current outputs are:

 - Mean Point A dose per patient
 - HR-CTV volume vs. HR-CTV D90%
 - Bladder volume vs. bladder D2cc
 - Bowel D2cc vs rectum D2cc
 - Fraction number vs. HR-CTV D90%
 - HR-CTV volume vs. fraction number


## Built with

[pyexcel_xls](https://pypi.python.org/pypi/pyexcel-xls/0.3.0)

[pymongo](https://api.mongodb.com/python/current/)

[ghostdriver](https://github.com/detro/ghostdriver)

[selenium](http://selenium-python.readthedocs.io/)

[bokeh](http://bokeh.pydata.org/en/latest/)
