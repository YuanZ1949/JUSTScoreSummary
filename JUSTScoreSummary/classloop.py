
import JUSTScoreSummary as jss
from os import system

file = open('account.txt','r')

a = file.readlines()
#print(a)
for i in a:
    account = i.split(' ')[0]
    password = i.split(' ')[1].replace('\n','')
    print(account, len(account))
    print(password , len(password))
    print()
    jss.main(eval(account),password)
    print('-'*20)
    system('pause')