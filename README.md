# <img src="https://github.com/iqp-samples/iqp-samples-ws/raw/master/logo.png" alt="iQuipsys Logo" width="100px" height="100px"> Excel integration sample

This is an excel integraion sample for [IQuipsys Positron](http://www.iquipsys.com).
This sample allows you to see the report about object presence in site zones. 

## Installation

- Install [Power query](https://www.microsoft.com/en-US/download/details.aspx?id=39379) excell add-in. For Excel 2016 instalation power query doesn't needed.

- Clone this repository to local disk
```bash
> git clone https://github.com/iqp-samples/iqp-samples-integration-excel.git
```

## Usage

- Excel file structure

<img src="https://github.com/iqp-samples/iqp-samples-integration-excel/blob/master/img/excel_structure.jpg?raw=true" alt="file structure"> 

- Set the login and password values in the auth table

- Set the report parameters in params table. Hours and minutes has to be double digits, like "00", "07", "15". Note that last hour statistic included to report.

- Click **Load data** button to create report

## Results

- The result is a pivot table on same spreadsheet as a parameters for report. 
<img src="https://github.com/iqp-samples/iqp-samples-integration-excel/blob/master/img/report_example.jpg?raw=true" alt="report example"> 