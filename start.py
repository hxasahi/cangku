# -*- coding: cp936 -*-
'''   ����������
      ʱ�䣺2020.2.15 19:30
      �汾���ֿ����ϵͳv1.0   '''
      #��Ŀ״̬��δ��

from openpyxl import load_workbook
wb = load_workbook('F:\pyadmin\products.xlsx')
ws = wb.active
#��
print('')
print('������ID��ѯ��')
shangpin = input()
nums = ws.max_row
for i in range(2,nums+1):
    if shangpin == ws.cell(row=i,column=1).value:
        print('���ƣ�'),
        print(ws.cell(row=i,column=2).value)
        print('�۸�{}'.format(ws.cell(row=i,column=3).value))
#��
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
