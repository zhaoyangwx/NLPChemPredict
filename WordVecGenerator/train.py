import _locale
_locale._getdefaultlocale = (lambda *args: ['zh_CN', 'utf8'])

import gensim
import fasttext
from gensim.test.utils import common_texts
from gensim.models.keyedvectors import FastTextKeyedVectors
from gensim.models._utils_any2vec import compute_ngrams, ft_hash_broken

fi=open("school_unsupervised.txt","r")
str1=fi.readlines()

sentences=[]
for s in str1:
    t=[s.split(" ")]
    sentences=sentences+t

min_ngrams, max_ngrams = 3, 4
sum_ngrams = 0

fv=open("school_unsupervised_segment.txt","w")

for s in sentences:
    for w in s:
        ret = compute_ngrams(w, min_ngrams, max_ngrams)  
        for sr in ret:
            fv.write(sr)
            fv.write(" ")
        sum_ngrams += len(ret)
    fv.write("\n")
fv.close()

model=fasttext.train_unsupervised(input='train_unsupervised.txt', model='skipgram',minCount=1, minn=3, maxn=4,word_ngrams = 2,dim=100, epoch = 500000)
model.save_model('model_unsupervised.bin')

def savevectors(model,fpath):
    words=model.get_words()
    file_out1=open(fpath,'w')
    # the first line must contain number of total words and vector dimension
    file_out1.write(str(len(words)) + " " + str(model.get_dimension()) + "\n")

    # line by line, you append vectors to VEC file
    for w in words:
        v = model.get_word_vector(w)
        vstr = ""
        for vi in v:
            vstr += " " + str(vi)
        try:
            file_out1.write(w + vstr+'\n')
        except:
            pass
    file_out1.close()

def drawvectors(model,fpath):
    from matplotlib import pylab
    pylab.figure(figsize=(15,15))
    words=model.get_words()
    i=0
    for w in words:
        i+=1
        v = model.get_word_vector(w)
        pylab.scatter(v[0],v[1])
        pylab.annotate(w,xy=(v[0],v[1]))
        if i>1000: break
    pylab.show()

def savesentencevectors(model,fpath):
    words=model.get_words()
    file_out1=open(fpath,'w')
    # the first line must contain number of total words and vector dimension
    file_out1.write(str(len(str1)) + " " + str(model.get_dimension()) + "\n")

    # line by line, you append vectors to VEC file
    for sentstr in str1:
        sentstr=sentstr.replace("\n","")
        sentv=model.get_sentence_vector(sentstr)
        vstr=""
        for vi in sentv:
            vstr+=" " + str(vi)
        try:
            file_out1.write(sentstr + vstr+'\n')
        except:
            pass
    file_out1.close()

def drawsentencevectors(model,fpath):
    from matplotlib import pylab
    pylab.figure(figsize=(15,15))
    i=0
    for sentstr in str1:
        sentstr=sentstr.replace("\n","")
        sentv=model.get_sentence_vector(sentstr)
        pylab.scatter(sentv[0],sentv[1])
        pylab.annotate(sentstr,xy=(sentv[0],sentv[1]))
        if i>200: break
        i+=1
    pylab.show()

savevectors(model,'output_unsuperised_wordvector.vec')
savesentencevectors(model,'output_unsupervised_sentencevector.vec')
drawsentencevectors(model,'output_unsupervised_sentencevector.vec')

fo=open("school_unsupervised_words.txt","w")
for s in model.words:
    fo.write(s)
    fo.write("\n")
fo.close()

model_supervised=fasttext.train_supervised(input='train_supervised.txt',pretrainedVectors='output_unsuperised_wordvector.vec',word_ngrams = 2,dim=100, epoch = 50000)
fo=open("output_supervised_words.txt","w")
for s in model_supervised.words:
    fo.write(s)
    fo.write("\n")
fo.close()
model_supervised.save_model("model_supervised.bin")
savevectors(model_supervised,'output_superised_wordvector.vec')
savesentencevectors(model_supervised,'output_supervised_sentencevector.vec')
drawsentencevectors(model_supervised,'output_superised_wordvector.vec')
savelabels(model_supervised,'output_superised_labels.txt')
