def LoadVec(filename):
    fh = open(filename, "r")
    s = fh.readlines()
    title = s[0].replace("\n", "").split(" ")
    lines, dims = int(title[0]), int(title[1])
    words = []
    vects = []
    for i in range(1, lines + 1):
        cl = s[i].split(" ")
        words.append(" ".join(cl[0:len(cl) - dims]))
        vects.append([float(xx) for xx in cl[-dims:]])
    fh.close()
    return words, vects

def LoadVal(filename):
    fh = open(filename, "r")
    s = fh.readlines()
    fh.close()
    values = []
    for t in s:
        values.append(float(t))
    return values

def dist(a, b):
    d = 0
    for i in range(len(a)):
        d = d + (b[i] - a[i]) ** 2
    d = d ** 0.5
    return d

def vectadd(a, b):
    for i in range(len(a)):
        a[i] = a[i] + b[i]
    return a

def vectsub(a, b):
    for i in range(len(a)):
        a[i] = a[i] - b[i]
    return a

def vectdiv(a, b):
    if b == 0:
        return a
    for i in range(len(a)):
        a[i] /= b
    return a

def clonevect(a):
    b = [0] * len(a)
    for i in range(len(a)):
        b[i] = a[i]
    return b

maxdist = 9999
def iter(center, vect):
    # center id for each vector:
    clist = [0] * len(vect)

    # average location (mass center) of vectors surrounding each center
    avgloc = [[0] * len(vect[0])] * len(center)

    # vector count for each center
    ccount = [0] * len(center)

    for i in range(len(vect)):
        mindist = maxdist

        # search nearest center for vector i
        for j in range(len(center)):
            d = dist(vect[i], center[j])
            if d < mindist:
                mindist = d
                clist[i] = j

        # update average location for the center of vactor i
        vectadd(avgloc[clist[i]], vect[i])
        ccount[clist[i]] += 1

    distlist = [0] * len(center)
    # update center and calc movement
    for i in range(len(center)):
        if ccount[i] > 0:
            vectdiv(avgloc[i], ccount[i])
            distlist[i] = dist(avgloc[i], center[i])
            #center[i] = avgloc[i]
            vectdiv(vectadd(center[i], avgloc[i]), 2)
        else:
            newcenter = [0] * len(center[0])
            for j in range(len(center)):
                vectadd(newcenter, center[j])
            vectsub(newcenter, center[i])
            vectdiv(newcenter, len(center)-1)
            distlist[i] = dist(newcenter, center[i])
            #center[i] = newcenter
            vectdiv(vectadd(center[i], newcenter), 2)
    return distlist, clist


numcenter = 4
maxiter = 10
center = []
# words, vects = LoadVec('split\\CHO\\CHO_full.vec')
words, vects = LoadVec('sch1.vec')
# thermalval = LoadVal('split\\CHO\\CHO_enthalpy_full.txt')
thermalval = LoadVal('schid1.txt')

# generate division list
import random
dlist = []
totallen = 0
for i in range(numcenter+1):
    len1 = random.randint(0, len(vects))
    dlist.append(len1)
    totallen += len1

# initiate center
sumlen = 0
for i in range(numcenter):
    sumlen += dlist[i]
    center.append(clonevect(vects[int(sumlen * len(vects) / totallen)]))

distnow = 9999 * 0

while distnow > 1e-3:
    dlist1, clist = iter(center, vects)
    distnow = max(dlist1)
    print(dlist1)
    maxiter -= 1
    if maxiter <= 0:
        break

import matplotlib.pyplot as plt
from sklearn import metrics
from sklearn.cluster import KMeans, MiniBatchKMeans
clist = KMeans(n_clusters=numcenter, random_state=9).fit_predict(vects)

# print("\nResult:\n")
# for i in range(numcenter):
#     print(center[i])
#     print("\n")
#
# dlist1, clist = iter(center, vects)
# print(clist)

# save file
pcount = [0] * numcenter
for i in range(len(clist)):
    pcount[clist[i]] += 1

fh = []
lh = []
import time
timestr = time.strftime("%Y%m%d%H%M%S", time.localtime())
for i in range(numcenter):
    if pcount[i] == 0:
        fh.append(None)
        lh.append(None)
    else:
        fh.append(open("result_" + timestr + "_class_" + str(i) + ".txt", "w"))
        lh.append(open("resultthermal_" + timestr + "_class_" + str(i) + ".txt", "w"))
        fh[i].write(str(pcount[i]))
        fh[i].write(" ")
        fh[i].write(str(len(vects[0])))
        fh[i].write("\n")
for i in range(len(vects)):
    fh[clist[i]].write(words[i])
    for j in range(len(vects[0])):
        fh[clist[i]].write(" ")
        fh[clist[i]].write(str(vects[i][j]))
    fh[clist[i]].write("\n")
    lh[clist[i]].write(str(thermalval[i]))
    lh[clist[i]].write("\n")
for i in range(numcenter):
    if pcount[i] != 0:
        fh[i].close()
        lh[i].close()

