"""
calc_tax関数は本体価格と消費税率を引数として
受け取り、消費税額を返します
"""
def calc_tax(price,rate):
    tax = price * rate / 100
    #int関数を使って整数を返す
    return int(tax)
    

a = calc_tax(1249,10)
print(a)