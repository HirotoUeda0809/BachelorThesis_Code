########################################################################
#作成日: 2020/1/1 作成者: Hiroto Ueda(hrt.ueda0809@gmail.com)
#最終更新日: 2020/1/16
#計算用メインプログラム
########################################################################

#ライブラリの読み込み
import pulp
import time
import RAP_OJT as rp


#パラメータ情報の読み込み
#スタッフ系の情報の読み取り
A = rp.AllMembersList()
#各スタッフの契約日数
O = rp.StaffContract()
#日付系の情報の読み取り
N = rp.Days()
Month = rp.Month()
#S系(シフトに関する情報)
S = rp.Shifts()
Stime = rp.ShiftsTime()
low = rp.ShiftBound_ts("\\{:}月希望.xlsx".format(Month), "L")
up = rp.ShiftBound_ts("\\{:}月希望.xlsx".format(Month),"U")
#研修系読み取り
tas = rp.tas()
#スキル系の行列
SA = rp.SkillArray()
TSA = rp.TeachSkillArray()
TPA = rp.TPlanArray()
RFA = rp.RefreshArray("\\{:}月希望.xlsx".format(Month))


#問題宣言
problem = pulp.LpProblem("Staff-Scheduling", pulp.LpMinimize)


#変数宣言
y = {}
lam = {}
x = {}
q = {}
for a in A:
    for j in N:
        for s in S:
            y[a,j,s] = pulp.LpVariable("y({:},{:},{:})".format(a,j,s), cat = "Binary")
            lam[a,j,s] = pulp.LpVariable("lam({:},{:},{:})".format(a,j,s), cat = "Binary")
            x[a,j,s] = pulp.LpVariable("x({:},{:},{:})".format(a,j,s), cat = "Binary")
            q[a,j,s] = pulp.LpVariable("q({:},{:},{:})".format(a,j,s), lowBound = 0)
d = {}
for a in A:
    for s in S:
        d[a,s] = pulp.LpVariable("d({:},{:})".format(a,s), lowBound = 0)
g = {}
for j in N:
    for s in S:
        g[j,s] = pulp.LpVariable("g({:},{:})".format(j,s), lowBound = 0)
el = {}
eu = {}
p = {}
for a in A:
    el[a] = pulp.LpVariable("el({:})".format(a), lowBound = 0)
    eu[a] = pulp.LpVariable("eu({:})".format(a), lowBound = 0)
    p[a] = pulp.LpVariable("p({:})".format(a), lowBound = 0)


#定式化
#目的関数
w1 = 10
w2 = 10
w3 = 3
#w4 = 10
w4 = {1:10, 2:10, 3:10, 21:10, 4:10, 5:10, 6:10, 22:10, 23:10, 7:10, 24:10}
w5 = 1
#obj1:なるべく早い期間で研修を終わらせてほしい！
obj1 = w1 * sum(sum(sum((1 - lam[a,j,s]) for s in S) for j in N) for a in A)
#obj2:できるだけ研修スタッフは月内にシフトをこなしてもらう．
obj2 = w2 * sum(sum(d[a,s] for s in S) for a in A)
#obj3:月の労働日数を平均化するよ！
obj3 = w3 * sum((el[a]) for a in A)
#obj4:シフトの下限制約が考慮制約の時に使用
obj4 = sum(sum(w4[s] * g[j,s] for s in S) for j in N)
#obj4:スタッフの休暇希望が考慮制約のときに使用
#obj4 = w4 * sum(p[a] for a in A)
#obj5:できるだけ指導できるスタッフが入ってね！
obj5 = w5 * sum(sum(sum(q[a,j,s] for s in S) for j in N) for a in A)
#(15)
problem += obj1 + obj2 + obj3 + obj4 + obj5, 'Total Cost'

#制約条件
#(1)
for a in A:
    for j in N:
        for s in S:
            problem += x[a,j,s] <= lam[a,j,s]
#(2)
for a in A:
    for j in N:
        for s in S:
            problem += SA[a,s] <= lam[a,j,s]
#(3)
for a in A:
    for j in N:
        for s in S:
            if (SA[a,s] == 0):
                problem += lam[a,j,s] <= TPA[a,s]
#(4)
for a in A:
    for s in S:
        if (TPA[a,s] == 0):
            problem += sum(y[a,j,s] for j in N) <= 0
#(5)
for j in N:
    for s in S:
        problem += sum(y[a,j,s] for a in A) <= 1
#(6)
for a in A:
    for s in S:
        problem += tas[a,s] - d[a,s] <= sum(y[a,j,s] for j in N)
        problem += sum(y[a,j,s] for j in N) <= tas[a,s]
#(7)
for a in A:
    for j in N:
        for s in S:
            problem += tas[a,s] * lam[a,j,s] <= sum(TPA[a,s] * y[a,k,s] for k in range(1, j + 1))
#(8)
for a in A:
    for j in N:
        problem += sum(x[a,j,s] for s in S) + sum(y[a,j,s] for s in S) <= 1
#(9)
for j in N:
    for s in S:
        problem += sum(x[a,j,s] for a in A) + g[j,s] >= low[j,s]
        #絶対制約切り替え時
        #problem += sum(x[a,j,s] for a in A) >= low[j,s]
        problem += sum(x[a,j,s] for a in A) <= up[j,s] 

#(10)
for a in A:
    problem += O[a] - el[a] <= sum(sum(x[a,j,s] for s in S) for j in N) + sum(sum(y[a,j,s] for s in S) for j in N)
    problem += O[a] + eu[a] >= sum(sum(x[a,j,s] for s in S) for j in N) + sum(sum(y[a,j,s] for s in S) for j in N)
#(11)
for a in A:
    for j in N:
        for (hopeS,realS) in zip(S,S):
            if (RFA[a,j] == hopeS):
                problem += sum((x[a,j,realS] + y[a,j,realS]) for realS in S if Stime[realS] < Stime[hopeS]) <= 0
#(12)
for a in A:
    problem += sum(RFA[a,j] * sum(x[a,j,s] + y[a,j,s] for s in S) for j in N if RFA[a,j] == 10) <= 0
#(13)
for b in A:
    for j in N:
        for s in S:
            problem += TPA[b,s] * y[b,j,s] <= sum(TSA[a,s] * x[a,j,s] for a in A) + q[b,j,s]


#求解パート
#問題の表示
print(problem)

start_time = time.perf_counter()
solver = problem.solve()
#CPLEX切り替え時に使用
#solver = problem.solve(pulp.CPLEX())
print("計算時間: {:.4f}秒".format(time.perf_counter() - start_time))

#最適性チェック
print(pulp.LpStatus[problem.status])

#計算結果を他ファイルに渡す
def optimal_solution(x_or_y_or_g):
    if (x_or_y_or_g == "x"):
        return x
    elif (x_or_y_or_g == "y"):
        return y
    elif (x_or_y_or_g == "g"):
        return g