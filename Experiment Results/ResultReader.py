import openpyxl
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import normaltest
from scipy.stats import mannwhitneyu

FunnyPath = "Experiment Results\Funny Results.xlsx"
SarcasticPath = "Experiment Results\Sarcastic Results.xlsx"

def cohens_d(x, y):
    nx = len(x)
    ny = len(y)
    dof = nx + ny - 2
    pooled_std = np.sqrt(((nx-1)*np.var(x, ddof=1) + (ny-1)*np.var(y, ddof=1)) / dof)
    d = (np.mean(x) - np.mean(y)) / pooled_std
    return abs(d)


# note that you cannot read the file while the file is opened with excel
FunnyWorkbook = openpyxl.load_workbook(FunnyPath)
SarcasticWorkbook = openpyxl.load_workbook(SarcasticPath)


FunnySheet = FunnyWorkbook.active
SarcasticSheet = SarcasticWorkbook.active

HesitationColumn = "V"
AnthroColumn = ["X","Y","Z","AA","AB"]
AnimacyColumn = ["AC",'AD','AE','AF','AG']
LikeabilityColumn = ['AH','AI','AJ','AK','AL']


# Read datas
FunnyHesitation = [cell.value for cell  in FunnySheet[HesitationColumn] if isinstance(cell.value, (int,float))]

SarcasticHesitation = [cell.value for cell  in SarcasticSheet[HesitationColumn]  if isinstance(cell.value, (int,float))]


# note that the first element in the arrays below is the question title
FunnyAnthro = []
SarcasticAnthro = []
for _ in AnthroColumn:
    Ftemp = [cell.value for cell  in FunnySheet[_] ]
    Ftemp.pop(0)
    Stemp = [cell.value for cell  in SarcasticSheet[_] ]
    Stemp.pop(0)
    
    FunnyAnthro.append(Ftemp)
    SarcasticAnthro.append(Stemp)


FunnyAnthro = np.array(FunnyAnthro).astype(float)
SarcasticAnthro  = np.array(SarcasticAnthro).astype(float)


FunnyAnimacy = []
SarcasticAnimacy = []
for _ in AnimacyColumn:
    Ftemp = [cell.value for cell  in FunnySheet[_] ]
    Ftemp.pop(0)
    Stemp = [cell.value for cell  in SarcasticSheet[_] ]
    Stemp.pop(0)
    
    FunnyAnimacy.append(Ftemp)
    SarcasticAnimacy.append(Stemp)

FunnyAnimacy = np.array(FunnyAnimacy).astype(float)
SarcasticAnimacy = np.array(SarcasticAnimacy).astype(float)


FunnyLike = []
SarcasticLike = []
for _ in LikeabilityColumn:
    Ftemp = [cell.value for cell  in FunnySheet[_] ]
    Ftemp.pop(0)
    Stemp = [cell.value for cell  in SarcasticSheet[_] ]
    Stemp.pop(0)
    
    FunnyLike.append(Ftemp)
    SarcasticLike.append(Stemp)

FunnyLike = np.array(FunnyLike).astype(float)
SarcasticLike = np.array(SarcasticLike).astype(float)

########################################################################

# visualisation

# Hesitation
'''
HesMean = [np.mean(FunnyHesitation),np.mean(SarcasticHesitation)]
HesSD = [np.std(FunnyHesitation),np.std(SarcasticHesitation)]

bar_width = 0.5
opacity = 0.8

fig,ax = plt.subplots()
rects = ax.bar(np.arange(len(HesMean)), HesMean, bar_width,
               alpha=opacity, color='b', yerr=HesSD, capsize=5)

ax.set_title('Hesitation of shutting down the robot')
ax.set_xlabel('Groups')
ax.set_ylabel('Hesitation(0-10), 0 means no hesitation')

ax.set_xticks(np.arange(len(HesMean)))
ax.set_xticklabels(('Funny', 'Sarcastic'))
#plt.show()

_,pvalue = mannwhitneyu(FunnyHesitation, SarcasticHesitation)

print(pvalue)

d = cohens_d(FunnyHesitation, SarcasticHesitation)

print(d)
'''

# Anthro
'''
FunnyAnthroMean = []
FunnyAnthroSD = []
SarcAnthroMean=[]
SarcAnthroSD = []

for i in FunnyAnthro:
    FunnyAnthroMean.append(np.mean(i))
    FunnyAnthroSD.append(np.std(i))
for i in SarcasticAnthro:
    SarcAnthroMean.append(np.mean(i))
    SarcAnthroSD.append(np.std(SarcAnthroSD))

bar_width = 0.35
x_axis = np.arange(len(FunnyAnthroMean))

color1 = "orange"
color2 = "blue"
plt.bar(x_axis, FunnyAnthroMean, width=bar_width, label='Funny', color=color1,yerr = FunnyAnthroSD, align="center",ecolor="black",capsize=4)
plt.bar(x_axis + bar_width, SarcAnthroMean, width=bar_width, label='Sarcastic', color=color2,yerr = FunnyAnthroSD, align="center",ecolor="black",capsize=4)

plt.legend()
plt.ylabel("Rating (1 for the former, 5 for the latter)")
plt.title("Anthropomorphism Meter")
plt.ylim(1,None)

labels = ["Fake-Natural","Machinelike-Humanlike","Unconscious-Conscious","Artificial-Lifelike","Rigidly-Elegantly"]
plt.xticks(np.arange(len(FunnyAnthroMean)),labels,rotation=10,fontsize = 8)
plt.show()

for i in range(len(FunnyAnthro)):
    _,p = mannwhitneyu(FunnyAnthro[i], SarcasticAnthro[i])
    d = cohens_d(FunnyAnthro[i], SarcasticAnthro[i])

    print(labels[i])
    print("P = "+str(p))
    print("d = "+str(d))
'''

'''
# Animacy

FunnyAnimacyMean = []
FunnyAnimacySD = []
SarcAnimacyMean=[]
SarcAnimacySD = []

for i in FunnyAnimacy:
    FunnyAnimacyMean.append(np.mean(i))
    FunnyAnimacySD.append(np.std(i))
for i in SarcasticAnimacy:
    SarcAnimacyMean.append(np.mean(i))
    SarcAnimacySD.append(np.std(SarcAnimacySD))

bar_width = 0.35
x_axis = np.arange(len(FunnyAnimacyMean))

color1 = "orange"
color2 = "blue"
plt.bar(x_axis, FunnyAnimacyMean, width=bar_width, label='Funny', color=color1,yerr = FunnyAnimacySD, align="center",ecolor="black",capsize=4)
plt.bar(x_axis + bar_width, SarcAnimacyMean, width=bar_width, label='Sarcastic', color=color2,yerr = FunnyAnimacySD, align="center",ecolor="black",capsize=4)

plt.legend()
plt.ylabel("Rating (1 for the former, 5 for the latter)")
plt.title("Animacy Meter")
plt.ylim(1,None)

labels = ["Dead-Alive","Stagnant-Lively","Mechanical-Organic","Inert-Interactive","Apathetic-Responsive"]
plt.xticks(np.arange(len(FunnyAnimacyMean)),labels,rotation=10,fontsize = 8)
#plt.show()

for i in range(len(FunnyAnimacy)):
    _,p = mannwhitneyu(FunnyAnimacy[i], SarcasticAnimacy[i])
    d = cohens_d(FunnyAnimacy[i], SarcasticAnimacy[i])

    p = round(p,3)
    d = round(d,3)

    print("\item "+labels[i]+"\\\\")
    print("P = "+str(p)+"\\\\")
    print("d = "+str(d))
'''
'''
# likeability


FunnyLikeMean = []
FunnyLikeSD = []
SarcLikeMean=[]
SarcLikeSD = []

for i in FunnyLike:
    FunnyLikeMean.append(np.mean(i))
    FunnyLikeSD.append(np.std(i))
for i in SarcasticLike:
    SarcLikeMean.append(np.mean(i))
    SarcLikeSD.append(np.std(SarcLikeSD))

bar_width = 0.35
x_axis = np.arange(len(FunnyLikeMean))

color1 = "orange"
color2 = "blue"
plt.bar(x_axis, FunnyLikeMean, width=bar_width, label='Funny', color=color1,yerr = FunnyLikeSD, align="center",ecolor="black",capsize=4)
plt.bar(x_axis + bar_width, SarcLikeMean, width=bar_width, label='Sarcastic', color=color2,yerr = FunnyLikeSD, align="center",ecolor="black",capsize=4)

plt.legend()
plt.ylabel("Rating (1 for the former, 5 for the latter)")
plt.title("Likeability Meter")
plt.ylim(1,None)

labels = ['Dislike-Like','Unfriendly-Friendly','Unkind-Kind','Unpleasant-Pleasant','Awful-Nice']
plt.xticks(np.arange(len(FunnyLikeMean)),labels,rotation=10,fontsize = 8)
plt.show()

for i in range(len(FunnyLike)):
    _,p = mannwhitneyu(FunnyLike[i], SarcasticLike[i])
    d = cohens_d(FunnyLike[i], SarcasticLike[i])

    p = round(p,3)
    d = round(d,3)

    print("\item "+labels[i]+"\\\\")
    print("P = "+str(p)+"\\\\")
    print("d = "+str(d))
'''