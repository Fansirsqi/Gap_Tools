a=['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k']
b = {_i:i for _i,i in enumerate(a,start=1)}
print(b)
if 'a' in b.values():
    print('ok')