import openpyxl as op

wb = op.load_workbook('example.xlsx')
snames = wb.sheetnames
s1 = wb['Hoja1']

save_rate = s1['A1'].value

if save_rate == None:
    save_rate = int(input("What percentage of your money would you like to save? "))
    save_rate = save_rate / 100
    s1['A1'] = save_rate


#def gain(n):

#def lose(n):

print("Welcome to the Personal Money Manager!\n")
wl = input("Did you gain or lose money:\n")
print("How much? ")
quantity = input()

if (wl == "gain"):
    print(save_rate)
    #gain(g)
else:
    print("sug")
    print(save_rate)
    #lose(n)

wb.save('example.xlsx')

