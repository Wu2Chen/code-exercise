#endcoding: utf-8

year = input("please input a year: ")

if (int(year) % 4 == 0 & int(year) % 100 == 0):
    if int(year) % 400 == 0:
        print("run year")
    else:
        print("not")
else:
    print("not")

