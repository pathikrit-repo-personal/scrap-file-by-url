import xlrd
import urllib.error
import urllib.parse
import urllib.request
import os
# Give the location of the file
loc = ("test.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
# print(sheet.nrows)
if not os.path.exists(
        os.path.realpath(os.path.dirname(os.path.dirname(__file__))) + "\product_images"):
    logText = {
        "message": "image directory not found creating one.",
        "title": "product_images",
        "path": os.path.dirname(os.path.dirname(__file__))
    }
    print(logText)
    os.makedirs(
        os.path.realpath(os.path.dirname(os.path.dirname(__file__))) + "\product_images")
if sheet.nrows > 0:
    for index in range(sheet.nrows):
        product_image_url=sheet.cell_value(index, 0) 
        if product_image_url:
            filename = product_image_url.split('/')[-1]
            fullfilename = os.path.join(
                os.path.realpath(
                    os.path.dirname(os.path.dirname(__file__))) + "\product_images",
                filename)
            try:
                urllib.request.urlretrieve(product_image_url, fullfilename)
            except urllib.error.HTTPError as e:
                logText = {
                    "message": "HTTPError Occurred while scrapping " ,
                    "url": product_image_url,
                    "error": 'HTTPError: {}'.format(e.reason)
                }
                print(logText)
            except urllib.error.URLError as e:
                logText = {
                    "message": "URLError Occurred while scrapping ",
                    "url": product_image_url,
                    "error": 'URLError: {}'.format(e.reason)
                }
                print(logText)
            else:
                logText = {
                    "message": "Retrieving image ",
                    "title": product_image_url,
                }
                print(logText)
        else:
            logText = {
                "message": "No Image found",
                "title": product_image_url,
            }
            print(logText)
