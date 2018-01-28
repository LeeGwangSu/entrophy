# -*- coding: utf-8 -*-
import xlrd
import math
file = 'a.xlsm'
path = 'C:\Users\Administrator\Desktop\_4-1\graduation\compare\kpop\\'
path+=file
workbook = xlrd.open_workbook(path)
#workbook = xlrd.open_workbook('C:\Users\Administrator\Desktop\ptt.xlsx')
worksheet = workbook.sheet_by_index(0) # 1번째 sheet 읽기(sheet1)
print file
num_rows = worksheet.nrows # 열 갯수
num_cols = worksheet.ncols # 행 갯수

""" ///////////////window size(N)바꿔가면서 다시 실행 N=6////////////////// """

N=6  # window크기, 즉 k의 크기
row_val=[]
c_val=[]

set=[0]*N
cnt_set=[1]*N
flag=[0]*300

a=[0]*pow(N,N)
q=0

"""a배열은 수열은 엔트로피 계산을 했을 때 순위를 결정하기 위해 기본값을 설정한다
예를 들면 k=3일때는 (1,1,1)부터 ... (3,3,3)까지 3의 3승 = 27개를 초기화한다"""
# window size 6이라면
for i in range(1,N+1):
    for j in range(1,N+1):
        for k in range(1,N+1):
            for m in range(1,N+1):
                for n in range(1,N+1):
                    for o in range(1,N+1):
                        a[q] = i * pow(10, N - 1) + j * pow(10, N - 2) + k * pow(10,N-3) + m * pow(10,N-4) + n * pow(10,N-5) +  o * pow(10,N-6)
                        q += 1


cnt=[0]*pow(N,N) # 각 순위의 갯수를 위한 배열

result=0
cnt_total=0
entropy=0.0
for row_num in range(num_rows):
    row_val.append(worksheet.row_values(row_num))



cnt_max=[0]*160

for i in range(144,159,1):
    for j in range(num_rows):
        if(row_val[j][0]==i):
            cnt_max[i]+=1

max_choice=143

for i in range(144,159,1):
    if(cnt_max[i] >= cnt_max[max_choice]):
        max_choice=i
print max_choice

max_temp=143

for i in range(144,159,1):
    if(cnt_max[i] >= cnt_max[max_temp] and i!=max_choice):
        max_temp=i
print max_temp


max_choice = max_temp



for i in range(num_rows):
    if(row_val[i][0]==max_choice):
        c_val.append(row_val[i+1][0])
    #c_val.append(row_val[i][0]) # 1열 read
c_rows=len(c_val)

#for i in range(c_rows):
    #print c_val[i],

for i in range(0,c_rows-(N-1),1):
    num=0
    for m in range(0,N):
        set[m] = c_val[i+m] # set배열은 배열 중 k 갯수만큼만 딴 값들을 각각 저장 ex(10 8 6)
    #for m in range(0, N):
        #print set[m]
    for m in range(0,N):
        cnt_set[m]=1 # cnt_set배열은 set배열의 순위를 위한 배열(1로 초기화)

    for i in range(0,N):
        for k in range(0,300):
            flag[k]=0
        """    k 갯수 값들 순위 결정"""
        for j in range(0,N):
            if(set[i]>set[j] and flag[int(set[j])]==0 ):
                cnt_set[i]+=1
                flag[int(set[j])]=1
    for j in range(0,N):
        num += cnt_set[j]*pow(10,N-j-1) # 순위의 값들(3 1 2)이런 숫자를 비교를 위해 312로 만들어준다
    for j in range(0,pow(N,N)):
        if(num==a[j]): # 값이 같으면
            cnt[j]+=1 # 각 순위의 카운트 값 증가
            cnt_total+=1

for j in range(0,pow(N,N)):
    if(cnt[j]!=0):
        entropy += (float(cnt[j])/float(cnt_total)) * math.log(float(cnt[j])/float(cnt_total),2)
entropy *= -1
print
print "N : ",N

for i in range(0,pow(N,N)):
    print cnt[i],
print
print cnt_total

print "entropy : ",entropy
print "-----------------------------------------------"

""" ///////////////window size(N)바꿔가면서 다시 실행 N=6////////////////// """

N=7  # window크기, 즉 k의 크기
row_val=[]
c_val=[]

set=[0]*N
cnt_set=[1]*N
flag=[0]*300

a=[0]*pow(N,N)
q=0
"""a배열은 수열은 엔트로피 계산을 했을 때 순위를 결정하기 위해 기본값을 설정한다
예를 들면 k=3일때는 (1,1,1)부터 ... (3,3,3)까지 3의 3승 = 27개를 초기화한다"""
# window size 7이라면
for i in range(1,N+1):
    for j in range(1,N+1):
        for k in range(1,N+1):
            for m in range(1,N+1):
                for n in range(1,N+1):
                    for o in range(1,N+1):
                        for w in range(1,N+1):
                            a[q] = i * pow(10, N - 1) + j * pow(10, N - 2) + k * pow(10,N-3) + m * pow(10,N-4) + n * pow(10,N-5) +  o * pow(10,N-6) + w* pow(10,N-7)
                            q += 1


cnt=[0]*pow(N,N) # 각 순위의 갯수를 위한 배열

result=0
cnt_total=0
entropy=0.0
for row_num in range(num_rows):
    row_val.append(worksheet.row_values(row_num))



cnt_max=[0]*160

for i in range(144,159,1):
    for j in range(num_rows):
        if(row_val[j][0]==i):
            cnt_max[i]+=1

max_choice=143

for i in range(144,159,1):
    if(cnt_max[i] >= cnt_max[max_choice]):
        max_choice=i
print max_choice

max_temp=143

for i in range(144,159,1):
    if(cnt_max[i] >= cnt_max[max_temp] and i!=max_choice):
        max_temp=i
print max_temp


max_choice = max_temp



for i in range(num_rows):
    if(row_val[i][0]==max_choice):
        c_val.append(row_val[i+1][0])
    #c_val.append(row_val[i][0]) # 1열 read
c_rows=len(c_val)

#for i in range(c_rows):
    #print c_val[i],

for i in range(0,c_rows-(N-1),1):
    num=0
    for m in range(0,N):
        set[m] = c_val[i+m] # set배열은 배열 중 k 갯수만큼만 딴 값들을 각각 저장 ex(10 8 6)
    #for m in range(0, N):
        #print set[m]
    for m in range(0,N):
        cnt_set[m]=1 # cnt_set배열은 set배열의 순위를 위한 배열(1로 초기화)

    for i in range(0,N):
        for k in range(0,300):
            flag[k]=0
        """    k 갯수 값들 순위 결정"""
        for j in range(0,N):
            if(set[i]>set[j] and flag[int(set[j])]==0 ):
                cnt_set[i]+=1
                flag[int(set[j])]=1
    for j in range(0,N):
        num += cnt_set[j]*pow(10,N-j-1) # 순위의 값들(3 1 2)이런 숫자를 비교를 위해 312로 만들어준다
    for j in range(0,pow(N,N)):
        if(num==a[j]): # 값이 같으면
            cnt[j]+=1 # 각 순위의 카운트 값 증가
            cnt_total+=1

for j in range(0,pow(N,N)):
    if(cnt[j]!=0):
        entropy += (float(cnt[j])/float(cnt_total)) * math.log(float(cnt[j])/float(cnt_total),2)
entropy *= -1
print
print "N : ",N

for i in range(0,pow(N,N)):
    print cnt[i],
print
print cnt_total

print "entropy : ",entropy
print "-----------------------------------------------"