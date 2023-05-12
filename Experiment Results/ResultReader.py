import openpyxl
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import normaltest
from scipy.stats import mannwhitneyu

FunnyPath = "Experiment Results\Funny Results.xlsx"
SarcasticPath = "Experiment Results\Sarcastic Results.xlsx"


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

SarcasticHesitation = [cell.value for cell  in SarcasticSheet["W"]  if isinstance(cell.value, (int,float))]


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

FunnyAnthro = np.array(FunnyAnthro)
SarcasticAnthro  = np.array(SarcasticAnthro)

FunnyAnimacy = []
SarcasticAnimacy = []
for _ in AnimacyColumn:
    Ftemp = [cell.value for cell  in FunnySheet[_] ]
    Ftemp.pop(0)
    Stemp = [cell.value for cell  in SarcasticSheet[_] ]
    Stemp.pop(0)
    
    FunnyAnimacy.append(Ftemp)
    SarcasticAnimacy.append(Stemp)

FunnyAnimacy = np.array(FunnyAnimacy)
SarcasticAnimacy = np.array(SarcasticAnimacy)


FunnyLike = []
SarcasticLike = []
for _ in LikeabilityColumn:
    Ftemp = [cell.value for cell  in FunnySheet[_] ]
    Ftemp.pop(0)
    Stemp = [cell.value for cell  in SarcasticSheet[_] ]
    Stemp.pop(0)
    
    FunnyLike.append(Ftemp)
    SarcasticLike.append(Stemp)

FunnyLike = np.array(FunnyLike)
SarcasticLike = np.array(SarcasticLike)

########################################################################

# visualisation

# Hesitation

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
plt.show()
