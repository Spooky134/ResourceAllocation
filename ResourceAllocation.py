from pandas import read_excel
from pandas import ExcelWriter
from pandas import DataFrame

path = 'data/input.xlsx'
df = read_excel(path, sheet_name='Sheet1')

COUNT_FACTORY = df.shape[1] - 1
COUNT_CONDITION = df.shape[0] + 1

# таблица f
source_table = []
for i in range(COUNT_FACTORY):
    source_table.append([0] + df['f' + str(i + 1) + '(x)'].tolist())

# таблица X-F
F, X = [], []
for i in range(COUNT_FACTORY):
    F.append([0] * COUNT_CONDITION)
    X.append([0] * COUNT_CONDITION)

#
for i in range(COUNT_CONDITION):
    F[COUNT_FACTORY - 1][i] = source_table[COUNT_FACTORY - 1][i]
    X[COUNT_FACTORY - 1][i] = i


results = []
for i in range(COUNT_FACTORY - 2, -1, -1):
    buf2 = []
    for S in range(1, COUNT_CONDITION):
        buf1 = []
        if i == 0:
            S = COUNT_CONDITION - 1
        for x in range(S + 1):
            a = S - x
            b = source_table[i][x]
            c = F[i + 1][a]
            s = round(b + c, 1)
            if s >= F[i][S]:
                X[i][S] = x
                F[i][S] = s
            buf1.append([x, a, b, c, s])
        buf2.append(buf1)
    results.append(buf2)


money = []
for i in range(COUNT_FACTORY):
    s = sum(money)
    money.append(X[i][COUNT_CONDITION - 1 - s])

weights = []
for i in range(COUNT_FACTORY):
    weights.append(source_table[i][money[i]])

for i in range(COUNT_FACTORY):
    print(str(i+1) + ' - ' + str(money[i]) + ' - ' + str(weights[i]))

print('summa = ' + str(sum(weights)))


# запись данных в файл
path = 'data/output.xlsx'
writer = ExcelWriter(path)

df = DataFrame()
df['S'] = list(range(COUNT_CONDITION))
for i in range(len(F)-1, -1, -1):
    df['x'+str(i+1)+'(S)'] = X[i]
    df['F'+str(i+1)+'(S)'] = F[i]

try:
    df.to_excel(writer, sheet_name='1', index=False)
    for i, el in enumerate(results):
        for j, res in enumerate(el):
            df_buf = DataFrame()
            df_buf['x'] = [x[0] for x in res]
            df_buf['S-x'] = [x[1] for x in res]
            df_buf['fi(xi)'] = [x[2] for x in res]
            df_buf['Fi+1(Si-xi)'] = [x[3] for x in res]
            df_buf['sum'] = [x[4] for x in res]
            df_buf.to_excel(writer, sheet_name='iteration' + str(i + 1) + '.' + str(j+1), index=False)
except Exception as ex:
    print(ex)
finally:
    writer.save()
