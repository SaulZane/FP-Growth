import pandas as pd
import xlwings as xw
from mlxtend.frequent_patterns import fpgrowth
from mlxtend.preprocessing import TransactionEncoder

app = xw.App(visible=False, add_book=False)
wb = app.books.open('excel3.xlsx')
sheet = wb.sheets('Sheet1')
data = sheet.range('A1').expand().value
# 转换为字符串
for i in range(len(data)):
    data[i] = [str(x) for x in data[i] if x]

wb.close()
app.quit()


te = TransactionEncoder()
tr_ary = te.fit_transform(data)
df = pd.DataFrame(tr_ary, columns=te.columns_)
frequent_itemsets: frozenset = fpgrowth(df, min_support=0.02, use_colnames=True)  # 设置最小支持度为0.02
# 输出所有support的值和频繁项集
print(len(frequent_itemsets))
for i in range(len(frequent_itemsets)):
     print(frequent_itemsets['support'][i], frequent_itemsets['itemsets'][i])
