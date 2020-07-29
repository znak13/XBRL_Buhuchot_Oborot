def f1():

    x = 100
    def f2():
        global y
        x = 200
        y = 333
    f2()
    print(x)

x = 111
y = 222

f1()
print(x)
print(y)

