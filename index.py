import requests
import xlsxwriter

workbook = xlsxwriter.Workbook("MyExcel.xlsx")
worksheet = workbook.add_worksheet()


response = requests.get("https://dummyjson.com/products")

data = response.json()

row = 1
col = 0

worksheet.write(0, 0, "Product id"), worksheet.write(0, 1, "Title"), worksheet.write(0, 1, "Product Description"), worksheet.write(0, 2, "Product Price"), worksheet.write(0, 3, "Product Discount %"), worksheet.write(0, 4, "Product Rating"), worksheet.write(0, 5, "Product Stock")


for index in (data['products']):
    # print("ID:", index["id"], "TITLE:", index["title"], "DESCRIPTION:", index["description"], "PRICE:", index["price"],
    #       "DISCOUNT PERCENT:", index["discountPercentage"], "RATING:", index["rating"], "STOCK:", index["stock"])
    worksheet.write(row, col, index["id"]), worksheet.write(row, col+1, index["title"]), worksheet.write(row, col+2, index["description"]), worksheet.write(row, col+3, index["price"]), worksheet.write(row, col+4, index["discountPercentage"]), worksheet.write(row, col+5, index["rating"]), worksheet.write(row, col+6, index["stock"])
    row += 1


workbook.close()
