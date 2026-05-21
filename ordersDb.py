import openpyxl as op 

workbook = op.Workbook()
sheet = workbook.active

sheet["A1"] = "Order ID"
sheet["B1"] = "Customer Name"
sheet["C1"] = "Product"
sheet["D1"] = "Quantity"
sheet["F1"] = "Price"
sheet["E1"] = "Total"

workbook.save("orderDB.xlsx")