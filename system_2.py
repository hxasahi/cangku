# -*- coding: cp936 -*-
'''   ����������
      ʱ�䣺2020.2.16 19:54
      �汾���ֿ����ϵͳv1.1   '''
      #��Ŀ״̬��δ��

from openpyxl import load_workbook
wb = load_workbook('F:\pyadmin\products.xlsx')
ws = wb.active
#��
def chaxun():
    print('')
    print('������ID��ѯ��')
    productid = input()
    nums = ws.max_row
    for i in range(2,nums+1):
        if productid == ws.cell(row=i,column=1).value:
            print('���ƣ�'),
            print(ws.cell(row=i,column=2).value)
            print('')
            print('�۸�{}'.format(ws.cell(row=i,column=3).value))
#��
def addproduct():
    tianjia=[]
    print('���ID:')
    zancun=int(input())
    tianjia.append(zancun)
    print('�������')
    zancun=raw_input()
    tianjia.append(zancun)
    print('��Ӽ۸�')
    zancun=int(input())
    tianjia.append(zancun)
    ws.append(tianjia)
    wb.save('F:\pyadmin\products.xlsx')
#ɾ
def delete():
    print('')
    shanchuid=int(input('������ɾ����ƷID:'))
    shan=['','','']
    nums = ws.max_row
    for i in range(2,nums+1):
        if shanchuid == ws.cell(row=i,column=1).value and i != nums :
            ws.cell(row=i,column=1).value=ws.cell(row=nums,column=1).value
            ws.cell(row=i,column=2).value=ws.cell(row=nums,column=2).value
            ws.cell(row=i,column=3).value=ws.cell(row=nums,column=3).value
            ws.cell(row=nums,column=1).value=''
            ws.cell(row=nums,column=2).value=''
            ws.cell(row=nums,column=3).value=''
            wb.save('F:\pyadmin\products.xlsx')
        elif shanchuid == ws.cell(row=i,column=1).value and i == nums :
            ws.cell(row=i,column=1).value=''
            ws.cell(row=i,column=2).value=''
            ws.cell(row=i,column=3).value=''
            wb.save('F:\pyadmin\products.xlsx')
#��
def correct():
    print('')
    xiugaiid=int(input('�������޸���ƷID'))
    nums = ws.max_row
    for i in range(2,nums+1):
        if xiugaiid == ws.cell(row=i,column=1).value and i != nums :
            xiuid=int(input('�������޸ĺ���ƷID��'))
            ws.cell(row=i,column=1).value=xiuid
            xiuname=raw_input('�������޸ĺ���Ʒ���ƣ�')
            ws.cell(row=i,column=2).value=xiuname
            xiujiage=int(input('�������޸ĺ���Ʒ�۸�'))
            ws.cell(row=i,column=3).value=xiujiage
            wb.save('F:\pyadmin\products.xlsx')
print('')
print('')
print('                          ��ӭʹ�òֿ����ϵͳ')
print('')
print('1 ��ѯ        '),
print('2 ����      '),
print('3 ɾ��   '),
print('4 �޸�     ')
print('')
print('99 �˳�')
n = input('               �밴�¶�Ӧ����ѡ��')
while 1 :
    if n == 1 :
        chaxun()
    elif n == 2 :
        addproduct()
    elif n == 3 :
        delete()
    elif n == 4 :
        correct()
    elif n == 99 :
        break
    #n = input('               �밴�¶�Ӧ����ѡ��')
        
    else :
        print('')
        print('               ������Ϸ��ַ�!')
        
    print('')
    print('               1 ����    '),
    print('99 �˳�')
    print('')

    c = input('               �Ƿ������')
    
    while c != 1 and c != 99 :
        print('')
        print('               ������Ϸ��ַ�!')
        print('')
        print('               1 ����    '),
        print('99 �˳�')
        print('')
        c = input('               �Ƿ������')
    if c == 1 :
        print('')
        print('1 ��ѯ        '),
        print('2 ����      '),
        print('3 ɾ��   '),
        print('4 �޸�     ')
        print('')
        print('99 �˳�')
        print('')     
    elif c == 99 :
        break
    
    n = input('               �밴�¶�Ӧ����ѡ��')       
