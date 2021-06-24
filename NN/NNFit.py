import time
import torch
import math
import numpy as np
from torch import nn
from torch.nn import functional as F
from torch import optim
import torchvision
from matplotlib import pyplot as plt

from utils import plot_scatter

from pyqtgraph.Qt import QtGui, QtCore
import pyqtgraph as pg


def seed_torch(seed=0):
    np.os.environ['PYTHONHASHSEED'] = str(seed)
    np.random.seed(seed)
    torch.manual_seed(seed)
    torch.cuda.manual_seed(seed)
    torch.cuda.manual_seed_all(seed)  # if you are using multi-GPU.
    torch.backends.cudnn.benchmark = False
    torch.backends.cudnn.deterministic = True


seed = int(time.time())
seed_torch(seed)
print(seed)

batch_size = 20000
epoches = 3000
plot_intv = 50
lr_patience = 2000
lr_factor = 0.5
init_lr = 0.005
l_momentum = 0.79
n_hidden = 5000
torch.set_default_tensor_type(torch.FloatTensor)

minval = 0
maxval = 0
minvec = 0
maxvec = 0
def NormalizeVector(inputvec):
    global minvec
    global maxvec
    minvec = inputvec[0][0]
    maxvec = minvec
    for vec in inputvec:
        for item in vec:
            if minvec > item:
                minvec = item
            if maxvec < item:
                maxvec = item
    return [[(item - minvec) / (maxvec-minvec) for item in vec] for vec in inputvec]


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
    return words, torch.FloatTensor(NormalizeVector(vects))


def LoadVal(filename):
    fh = open(filename, "r")
    s = fh.readlines()
    fh.close()
    values = []
    for t in s:
        values.append([float(t)])
    values = Normalize(values)
    return torch.FloatTensor(values)



def Normalize(inputlist, Norm = True):
    global minval
    global maxval
    global meanval
    minval = min(min(inputlist))
    maxval = max(max(inputlist))
    meanval = np.average(inputlist)
    if minval == maxval or (not Norm):
        return inputlist
    return [[(x[0] - minval) / (maxval - minval) * 2 - 1] for x in inputlist]



def DeNorm(value, Denorm=True):
    if Denorm:
        return (value +1) / 2 * (maxval - minval) + minval
    return value


class WVDataset(torch.utils.data.Dataset):

    def __init__(self, dataset_type=None, transform=None, update_dataset=False):
        """
        dataset_type: ['train', 'test']
        """
        super(WVDataset, self).__init__()
        self.sample_list = list()
        self.label_list = list()
        self.word_list = list()
        self.dataset_type = dataset_type

    def __getitem__(self, index):
        item = self.sample_list[index]
        label = self.label_list[index]
        word = self.word_list[index]
        return item, label, word

    def __len__(self):
        return len(self.label_list)


words, vects = LoadVec('CHO_full.vec')
label = LoadVal('CHO_enthalpy_full.txt')


samplecount = vects.shape[0]
dims = vects.shape[1]
ds = WVDataset()
ds.label_list = label
ds.sample_list = vects
ds.word_list = words

train_db, val_db = torch.utils.data.random_split(ds, [samplecount - samplecount // 5, samplecount // 5])

train_loader = torch.utils.data.DataLoader(train_db, batch_size=batch_size, shuffle=True)
val_loader = torch.utils.data.DataLoader(val_db, batch_size=batch_size, shuffle=True)


class Net(nn.Module):
    def __init__(self):
        super(Net, self).__init__()

        self.fc1 = nn.Linear(dims, n_hidden)
        self.fc2 = nn.Linear(n_hidden, n_hidden)
        self.fc3 = nn.Linear(n_hidden, n_hidden)
        self.fcx = nn.Linear(n_hidden, 1)

    def forward(self, xx):
        xx = F.relu(self.fc1(xx))
        xx = F.relu(self.fc2(xx))
        xx = F.relu(self.fc3(xx))
        xx = self.fcx(xx)

        return xx


def validate(val_loader, net):
    __loss = 0
    __count = 0
    for __x, __y, __w in val_loader:
        __out = net(__x.cuda().half())
        # out: [b, 10] -> pred: [b]
        co = F.mse_loss(__out, __y.cuda().half())
        __loss += co
        __count += 1
    return __loss / __count


def logplotdataconv(x):
    return x
    if x == 0:
        return 0
    return math.log10(math.fabs(x))


def plot_eff(DeNormOption = False):
    val_trueval = []
    val_fitted = []
    x_trueval = []
    for batch_idx, (x, y, w) in enumerate(val_loader):
        x = x.cuda().half()
        y = y.cuda().half()
        for yn in y:
            val_trueval.append(logplotdataconv(DeNorm(yn.item(), DeNormOption)))
        for xn in x:
            xnorm = torch.mean(xn)
            val_fitted.append(logplotdataconv(DeNorm(net(xn).item(), DeNormOption)))
            x_trueval.append(logplotdataconv(xnorm.item()))
    return x_trueval, val_fitted, val_trueval

def plot_teff(DeNormOption = False):
    val_trueval = []
    val_fitted = []
    x_trueval = []
    for batch_idx, (x, y, w) in enumerate(train_loader):
        x = x.cuda().half()
        y = y.cuda().half()
        for yn in y:
            val_trueval.append(logplotdataconv(DeNorm(yn.item(), DeNormOption)))
        for xn in x:
            xnorm = torch.mean(xn)
            val_fitted.append(logplotdataconv(DeNorm(net(xn).item(), DeNormOption)))
            x_trueval.append(logplotdataconv(xnorm.item()))
    return x_trueval, val_fitted, val_trueval


net = Net()
net = net.cuda().half()
# [w1, b1, w2, b2, w3, b3]
optimizer = optim.SGD(net.parameters(), lr=init_lr, momentum=l_momentum, weight_decay=0.001)
scheduler = torch.optim.lr_scheduler.ReduceLROnPlateau(optimizer, 'min', factor=lr_factor, patience=lr_patience,
                                                       verbose='Reduced')
train_loss = []
t0 = time.time()
batch_idx = 0
print('Preparing PyQt')
trainflag = 0
app = pg.mkQApp()
pw = pg.plot()
pw.setBackground('w')
pw.showGrid(x=True, y=True)
pw.setXRange(min=0.47, max=0.53)
pw.setYRange(min=-0.45, max=1)

app2 = pg.mkQApp()
pw2 = pg.plot()
pw2.setBackground('w')
pw2.showGrid(x=True, y=True)
pw2.setXRange(min=0, max=20)
pw2.setYRange(min=0, max=0.02)
print('Init OK\n')

ax = []
aytrain = []
ayval = []

def trainthread():
    global trainflag
    for epoch in range(epoches):
        for batch_idx, (x, y, w) in enumerate(train_loader):
            x = x.cuda().half()
            y = y.cuda().half()
            out = net(x)
            loss = F.mse_loss(out, y)

            optimizer.zero_grad()
            loss.backward()
            optimizer.step()

        test_lossval = validate(val_loader=val_loader, net=net)
        train_lossval = validate(val_loader=train_loader, net=net)
        print("epoch{} trainloss={} valloss={}".format(epoch, train_lossval, test_lossval))
        ax.append(epoch)
        aytrain.append(train_lossval.item())
        ayval.append(test_lossval.item())
        scheduler.step(test_lossval)
        if epoch % plot_intv == 0:
            px, py, ptrue = plot_eff()
            pxt, pyt, ptruet = plot_teff()
            pw.clear()
            pw2.clear()
            st = pg.ScatterPlotItem(pen=pg.mkPen('g'))
            st.addPoints(x=ptruet, y=pyt)
            s0 = pg.ScatterPlotItem(pen=pg.mkPen('b'))
            s0.addPoints(x=ptrue, y=py)
            ltr = pg.PlotCurveItem(x=np.array(ax), y=np.array(aytrain), pen=pg.mkPen('g'))
            lval = pg.PlotCurveItem(x=np.array(ax), y=np.array(ayval), pen=pg.mkPen('b'))
            pw.addItem(st)
            pw.addItem(s0)
            pw.setXRange(min=-1, max=1)
            pw.setYRange(min=-1, max=1)
            pw2.setXRange(min=0, max=epoch+10)
            pw2.addItem(ltr)
            pw2.addItem(lval)
    trainflag = 1


import _thread

_thread.start_new_thread(trainthread, ())
QtGui.QApplication.instance().exec_()

while trainflag == 0:
    continue
t1 = time.time()
print("Total train time = ", t1 - t0)

acc = validate(val_loader=val_loader, net=net)
val_trueval = []
val_fitted = []
for batch_idx, (x, y, w) in enumerate(val_loader):
    for yn in y:
        val_trueval.append(logplotdataconv(DeNorm(yn.item())))
    for xn in x:
        val_fitted.append(logplotdataconv(DeNorm(net(xn.cuda().half()).item())))
plot_scatter(val_trueval, val_fitted, 'True Val', 'Fitted')
print('test l2err:{}'.format(acc))

train_molname = []
train_trueval = []
train_pred = []
for batch_idx, (x, y, w) in enumerate(train_loader):
    for yn in y:
        train_trueval.append(DeNorm(yn.item()))
    for xn in x:
        train_pred.append(DeNorm(net(xn.cuda().half()).item()))
    for wn in w:
        train_molname.append(wn)

val_molname = []
val_trueval = []
val_pred = []
for batch_idx, (x, y, w) in enumerate(val_loader):
    for yn in y:
        val_trueval.append(DeNorm(yn.item()))
    for xn in x:
        val_pred.append(DeNorm(net(xn.cuda().half()).item()))
    for wn in w:
        val_molname.append(wn)

fntstmp = time.strftime("%Y%m%d%H%M%S", time.localtime()) + "_" + str(seed)
fw = open("result_train_" + fntstmp + ".txt", "w")
for i in range(len(train_molname)):
    fw.write(train_molname[i])
    fw.write("\t")
    fw.write(str(train_trueval[i]))
    fw.write("\t")
    fw.write(str(train_pred[i]))
    fw.write("\n")
fw.close()


fw = open("result_validate_" + fntstmp + ".txt", "w")
for i in range(len(val_molname)):
    fw.write(val_molname[i])
    fw.write("\t")
    fw.write(str(val_trueval[i]))
    fw.write("\t")
    fw.write(str(val_pred[i]))
    fw.write("\n")
fw.close()

exit()

label_out = []
val_inplbl = []
val_pred = []
for i in range(len(vects)):
    label_out.append(words[i])
    val_inplbl.append(DeNorm(label[i].item()))
    val_pred.append(DeNorm(net(vects[i]).item()))
fw = open("result_" + time.strftime("%Y%m%d%H%M%S", time.localtime()) + ".txt", "w")
for i in range(len(val_inplbl)):
    fw.write(label_out[i])
    fw.write("\t")
    fw.write(str(val_inplbl[i]))
    fw.write("\t")
    fw.write(str(val_pred[i]))
    fw.write("\n")
fw.close()