<p align="center">
<img src="https://github.com/aaronphaneuf/flyer_performance/blob/master/images/flyer_performance.PNG">
</p>

# What is flyer_performance_build.py?

Compares movement data and writes it to a .xlsx file

## What is the purpose of flyer_performance_build.py?

Each month, my employer runs a circulated flyer of items on promotion. A report called Flyer Performance is generated to compare item movement before and after the flyer promotion. Additionally, a second sheet called Brand Subtotals is included which sums the total sales for each brand. Up until January 2020, this report was built manually
which is time consuming and can lead room for errors. The script I have put together pulls all the relevant data and presents it in exactly the same fashion, with absolute accuracy. The time to build has also been cut down from hours to seconds.

## Requirements

flyer_performance_build.py requires python 3.x and the following modules:

Pandas
Pyodbc

# Usage

To know what items are being compared, a file containing each item on promotion is compiled along with the following identifiers:

UPC Brand Description Size Promo Cost Promo Retail

Upon running the python file, the user is greeted with the following text:

Flyer Performance Build
1. Build Report
2. Debug Mode
Selection Option:

Selecting option "1" will ask for a month and year, corresponding to the dates the flyer are effective.
Previous movement is equal to flyer start date - the amount of days the flyer runs while future movement is equal
to the flyer end date + the amount of days the flyer runs. This gives the interested party a good idea of how well the item(s) sold before and after the flyer became effective.

Selecting option "2" allows the user without any knowledge of Python to change or add flyer dates.
