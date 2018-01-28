# -*- coding: utf-8 -*-
import xlrd
import math
import matplotlib.pyplot as plt

file = 'a.xlsm'
path = 'C:\Users\Administrator\Desktop\_4-1\graduation\compare\pop\\'
path+=file
workbook = xlrd.open_workbook(path)
#workbook = xlrd.open_workbook('C:\Users\Administrator\Desktop\ptt.xlsx')
worksheet = workbook.sheet_by_index(0) # 1번째 sheet 읽기(sheet1)
print file
num_rows = worksheet.nrows # 열 갯수
num_cols = worksheet.ncols # 행 갯수

#------------------------- 3 ----------------------
# ///////////환경설정
N=3  # 작은 window 크기, 즉 k의 크기
part_size=50 # 끊은 큰 window 크기

row_val=[]
c_val=[]

set=[0]*N
cnt_set=[1]*N
flag=[0]*300

a=[0]*pow(N,N)
q=0

for i in range(1,N+1):
    for j in range(1,N+1):
        for k in range(1,N+1):
            a[q]=i*pow(10,N-1) + j*pow(10,N-2) + k*pow(10,N-3)
            q+=1

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

max_choice=144

for i in range(144,159,1):
    if(cnt_max[i] >= cnt_max[max_choice]):
        max_choice=i

for i in range(num_rows):
    if(row_val[i][0]==max_choice): # 해당 악기 음(예 144)일 때
        c_val.append(row_val[i+1][0]) # 1열 read 다음 숫자(음계)
c_rows=len(c_val)
for i in range(c_rows):
    print c_val[i],
arr_part_entropy=[0]*c_rows
idx_arr=0



# 환경설정///////////
for i in range(part_size):
    print c_val[i],
#print
for z in range(0,c_rows-(part_size-1),1):
    #for u in range(z,z+part_size,1):
        #print c_val[u],
    part_result = 0
    part_cnt_total = 0
    part_entropy = 0.0

    part_set = [0] * N
    part_cnt_set = [1] * N
    part_cnt = [0] * pow(N, N)
    part_flag = [0] * 300

    for i in range(z,z+part_size-(N-1),1):
        num=0
        for m in range(0,N):
            part_set[m] = c_val[i + m]
        for m in range(0,N):
            part_cnt_set[m]=1
        for n in range(0,N):
            for k in range(0,300):
                part_flag[k]=0
            for j in range(0, N):
                if (part_set[n] > part_set[j] and part_flag[int(part_set[j])] == 0):
                    part_cnt_set[n] += 1
                    part_flag[int(part_set[j])] = 1
        for j in range(0, N):
            num += part_cnt_set[j] * pow(10, N - j - 1)  # 순위의 값들(3 1 2)이런 숫자를 비교를 위해 312로 만들어준다
        for j in range(0, pow(N, N)):
            if (num == a[j]):  # 값이 같으면
                part_cnt[j] += 1  # 각 순위의 카운트 값 증가
                part_cnt_total += 1

    for j in range(0, pow(N, N)):
        if (part_cnt[j] != 0):
            part_entropy += (float(part_cnt[j]) / float(part_cnt_total)) * math.log(float(part_cnt[j]) / float(part_cnt_total), 2)
    part_entropy *= -1

    """
    print
    print "N : ", N
    for i in range(0, pow(N, N)):
        print part_cnt[i],
    print
    print part_cnt_total
    print "part_entropy : ", part_entropy
    print "///////////////////////////////////////////////////////////"
    """
    arr_part_entropy[idx_arr]=part_entropy
    idx_arr+=1


plt.figure(figsize=(50,20))
x1= range(0,idx_arr)
y1=[arr_part_entropy[v] for v in x1]

plt.subplot(3,1,1)

plt.plot(x1,y1,'k-',label="k=3",color="red",linewidth=3)
plt.legend(loc='best')
plt.ylim([0,4])
plt.title('nowimhere-queen_50')
#------------------------- 4 ----------------------

N=4  # 작은 window 크기, 즉 k의 크기

# /////////// 환경설정
row_val=[]
c_val=[]

set=[0]*N
cnt_set=[1]*N
flag=[0]*300

a=[0]*pow(N,N)
q=0
# window size 4이라면
for i in range(1,N+1):
    for j in range(1,N+1):
        for k in range(1,N+1):
            for m in range(1,N+1):
                a[q] = i * pow(10, N - 1) + j * pow(10, N - 2) + k * pow(10,N-3) + m * pow(10,N-4)
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

max_choice=144

for i in range(144,159,1):
    if(cnt_max[i] >= cnt_max[max_choice]):
        max_choice=i

print max_choice

for i in range(num_rows):
    if(row_val[i][0]==max_choice): # 해당 악기 음(예 144)일 때
        c_val.append(row_val[i+1][0]) # 1열 read 다음 숫자(음계)
c_rows=len(c_val)

print c_rows
arr_part_entropy=[0]*c_rows
idx_arr=0

# 환경설정  ///////////
print
for i in range(part_size):
    print c_val[i],
print
for z in range(0,c_rows-(part_size-1),1):
    #print
    #for u in range(z,z+part_size,1):
        #print c_val[u],
    part_result = 0
    part_cnt_total = 0
    part_entropy = 0.0

    part_set = [0] * N
    part_cnt_set = [1] * N
    part_cnt = [0] * pow(N, N)
    part_flag = [0] * 300

    for i in range(z,z+part_size-(N-1),1):
        num=0
        for m in range(0,N):
            part_set[m] = c_val[i + m]
        for m in range(0,N):
            part_cnt_set[m]=1
        for n in range(0,N):
            for k in range(0,300):
                part_flag[k]=0
            for j in range(0, N):
                if (part_set[n] > part_set[j] and part_flag[int(part_set[j])] == 0):
                    part_cnt_set[n] += 1
                    part_flag[int(part_set[j])] = 1
        for j in range(0, N):
            num += part_cnt_set[j] * pow(10, N - j - 1)  # 순위의 값들(3 1 2)이런 숫자를 비교를 위해 312로 만들어준다
        for j in range(0, pow(N, N)):
            if (num == a[j]):  # 값이 같으면
                part_cnt[j] += 1  # 각 순위의 카운트 값 증가
                part_cnt_total += 1

    for j in range(0, pow(N, N)):
        if (part_cnt[j] != 0):
            part_entropy += (float(part_cnt[j]) / float(part_cnt_total)) * math.log(float(part_cnt[j]) / float(part_cnt_total), 2)
    part_entropy *= -1

    """print
    print "N : ", N
    for i in range(0, pow(N, N)):
        print part_cnt[i],
    print 
    print part_cnt_total
    print "part_entropy : ", part_entropy
    print "///////////////////////////////////////////////////////////"
"""
    arr_part_entropy[idx_arr] = part_entropy
    idx_arr += 1


x2= range(0,idx_arr)
y2=[arr_part_entropy[v] for v in x2]


plt.subplot(3,1,2)
plt.plot(x2,y2,'k-',label="k=4",color="green",linewidth=3)
plt.legend(loc='best')
plt.ylim([0,6])


#------------------------- 5 ----------------------



N=5  # 작은 window 크기, 즉 k의 크기

# /////////// 환경설정
row_val=[]
c_val=[]

set=[0]*N
cnt_set=[1]*N
flag=[0]*300
a=[0]*pow(N,N)
q=0
# window size 5이라면
for i in range(1,N+1):
    for j in range(1,N+1):
        for k in range(1,N+1):
            for m in range(1,N+1):
                for n in range(1,N+1):
                    a[q] = i * pow(10, N - 1) + j * pow(10, N - 2) + k * pow(10,N-3) + m * pow(10,N-4) + n * pow(10,N-5)
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

max_choice=144

for i in range(144,159,1):
    if(cnt_max[i] >= cnt_max[max_choice]):
        max_choice=i

print max_choice

for i in range(num_rows):
    if(row_val[i][0]==max_choice): # 해당 악기 음(예 144)일 때
        c_val.append(row_val[i+1][0]) # 1열 read 다음 숫자(음계)
c_rows=len(c_val)

print c_rows
arr_part_entropy=[0]*c_rows
idx_arr=0

# 환경설정  ///////////
print
for i in range(part_size):
    print c_val[i],
print
for z in range(0,c_rows-(part_size-1),1):
    """
    for u in range(z,z+part_size,1):
        print c_val[u],
    print
    """
    part_result = 0
    part_cnt_total = 0
    part_entropy = 0.0

    part_set = [0] * N
    part_cnt_set = [1] * N
    part_cnt = [0] * pow(N, N)
    part_flag = [0] * 300

    for i in range(z,z+part_size-(N-1),1):
        num=0
        for m in range(0,N):
            part_set[m] = c_val[i + m]
        for m in range(0,N):
            part_cnt_set[m]=1
        for n in range(0,N):
            for k in range(0,300):
                part_flag[k]=0
            for j in range(0, N):
                if (part_set[n] > part_set[j] and part_flag[int(part_set[j])] == 0):
                    part_cnt_set[n] += 1
                    part_flag[int(part_set[j])] = 1
        for j in range(0, N):
            num += part_cnt_set[j] * pow(10, N - j - 1)  # 순위의 값들(3 1 2)이런 숫자를 비교를 위해 312로 만들어준다
        for j in range(0, pow(N, N)):
            if (num == a[j]):  # 값이 같으면
                part_cnt[j] += 1  # 각 순위의 카운트 값 증가
                part_cnt_total += 1

    for j in range(0, pow(N, N)):
        if (part_cnt[j] != 0):
            part_entropy += (float(part_cnt[j]) / float(part_cnt_total)) * math.log(float(part_cnt[j]) / float(part_cnt_total), 2)
    part_entropy *= -1

    """print
    print "N : ", N
    for i in range(0, pow(N, N)):
        print part_cnt[i],
    print
    print part_cnt_total
    print "part_entropy : ", part_entropy
    """




    arr_part_entropy[idx_arr] = part_entropy
    idx_arr += 1


x3= range(0,idx_arr)
y3=[arr_part_entropy[v] for v in x3]


plt.subplot(3,1,3)
plt.plot(x3,y3,'k-',label="k=5",color="blue",linewidth=3)
plt.legend(loc='best')
plt.ylim([0,7])
plt.savefig("nowimhere-queen.png")
plt.show()