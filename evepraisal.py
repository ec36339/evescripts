"""
Convert JSON output from https://evepraisal.com/ into CSV data that can be pasted into Excel.
The script arranges the output in the same order as the input list of items,
so the price data can be easily combined with existing data already in the spreadsheet
(such as average prices, market volume, etc.)
It requires two input files:
- items.txt: List of items (one item name per line)
- prices.json: JSON output from https://evepraisal.com/
Usage:
1. Copy the item list from your spreadsheet (usually a column containing item names)
2. Paste the item list into items.txt
3. Paste the item list into https://evepraisal.com/
4. Export the data from https://evepraisal.com/ as JSON and paste it into prices.json
5. Run this script
6. Paste the output into your spreadsheet.
Note: If you misspelled an item name, then its buy and sell price will be listed as zero.
"""

import json


def make_price(price):
    """
    Convert an ISK price (float) to a format acceptable by Excel
    :param price: Float ISK price
    :return: ISK price as Excel string
    """
    return ('%s' % price).replace('.', ',')


def print_csv(name, buy, sell):
    """
    Print a row for an Excel sheet
    :param name: Name of item
    :param buy: Float buy order ISK price
    :param sell: Float sell order ISK price
    :return: ISK price as tab-separated Excel row
    """
    print("%s\t%s\t%s" % (name, make_price(buy), make_price(sell)))


def main():
    # Load list of items (copy column from Excel sheet)
    with open('items.txt') as f:
        items = [l.split('\n')[0] for l in f.readlines()]

    # Load price data (export JSON from evepraisal.com after pasting same item list)
    with open('prices.json') as f:
        j = json.load(f)

    # Arrange prices in lookup table by item name
    item_db = {i['name']: i['prices'] for i in j['items']}

    # Print item data as Excel row with items in same order as original item list.
    for i in items:
        found = item_db.get(i)
        if found is not None:
            print_csv(i, found['buy']['max'], found['sell']['min'])
        else:
            print_csv(i, 0, 0)


if __name__ == '__main__':
    main()
