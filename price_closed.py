import matplotlib.pyplot as plt
import pandas as pd
import win32com.client

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

instStockChart.SetInputValue(0, "A005380")
instStockChart.SetInputValue(1, ord('2'))
instStockChart.SetInputValue(4, 252)
instStockChart.SetInputValue(5, (0, 5))
instStockChart.SetInputValue(6, ord('D'))
instStockChart.SetInputValue(9, ord('1'))

instStockChart.BlockRequest()

numData = instStockChart.GetHeaderValue(3)
numField = instStockChart.GetHeaderValue(1)

dates = []
price = []

for j in range(numField):
    if j == 0:
        for i in range(numData):
            dates.append(instStockChart.GetDataValue(j, i))
    if j == 1:
        for i in range(numData):
            price.append(instStockChart.GetDataValue(j, i))

df = pd.DataFrame(price, index=pd.to_datetime(dates, format='%Y%m%d', errors='ignore'),
                  columns=['stock'])
df.index
df.plot()
plt.show()
