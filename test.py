def fn(x):
    return 3*x*x+2*x+1

def recur(a,r,num=1):
    
    if r==num:
        return fn(a)
    elif r==0:
        return 1
    else:
        return recur(fn(a),r,num+1)


# #50번 반복한 값
# print(recur(0,50))

#각 값들을 알고싶으면
num=15
for i in range(0,num+1):
    print("==================================")
    print("f"+str(i)+" = "+str(recur(0,i)))
    print("==================================")