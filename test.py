test_dict = {x:chr(x) for x in range(0,50)}
i=0
#print(test_dict)
for x in range(50,100):
    test_dict[i] = [test_dict[i],x]
    i+=1
#print(test_dict)

print(*enumerate(test_dict.values(),1))