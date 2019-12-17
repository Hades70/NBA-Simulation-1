#!/usr/bin/env python
# coding: utf-8

import pandas as pd
from plotly.offline import iplot, init_notebook_mode
from plotly.subplots import make_subplots
init_notebook_mode()
import plotly.graph_objs as go
import numpy as np
import random
import xlwings as xw

book = xw.Book(r'nba-data.csv')
sheet = book.sheets('Main')
df = sheet.range('a1:ag31').options(pd.DataFrame).value
df = df.drop(['Team'],axis=1)

dictDf = df.to_dict()

winPct = dictDf['win%']
pace = dictDf['pace']
tno = dictDf['to_rate']
orr = dictDf['off_rebr']
drr = dictDf['def_rebr']
rebr = dictDf['rebr']
offEff = dictDf['off_eff']
defEff = dictDf['def_eff']
netEff = dictDf['net_eff']
threePct = dictDf['3_pt%']
tsPct = dictDf['ts%']
oppTsPct = dictDf['opp_ts%']
netTsPct = dictDf['net_ts%']
threeRate = dictDf['3_rate']
effFg = dictDf['eff_fg%']
oppEffFg = dictDf['opp_eff_fg%']
netEffFg = dictDf['net_eff_fg%']
twoPct = dictDf['2_pt%']

class Team():
    def __init__(self,x,y):
        self.pace = pace[x]
        self.tno = tno[x]
        self.orr = orr[x]
        self.drr = drr[x]
        self.threePct = threePct[x]
        self.threeRate = threeRate[x]
        self.effFg = effFg[x]
        self.oppEffFg = oppEffFg[x]
        self.mascot = y
        self.name = x
        self.offEff = offEff[x]
        self.defEff = defEff[x]
        self.trueShooting = tsPct[x]
        self.oppTrueShooting = oppTsPct[x]
        self.winPct = winPct[x]
        self.rebRate = rebr[x]
        self.netTrueShooting = netTsPct[x]
        self.effectiveFg = effFg[x]
        self.oppEffectiveFg = oppEffFg[x]
        self.netEffectiveFg = netEffFg[x]
        self.twoPct = twoPct[x]
        
ATL = Team('Atlanta','Hawks')
BOS = Team('Boston','Celtics')
BKN = Team('Brooklyn','Nets')
CHA = Team('Charlotte','Hornets')
CHI = Team('Chicago','Bulls')
CLE = Team('Cleveland','Cavaliers')
DAL = Team('Dallas','Mavericks')
DEN = Team('Denver','Nuggets')
DET = Team('Detroit','Pistons')
GSW = Team('Golden State','Warriors')
HOU = Team('Houston','Rockets')
IND = Team('Indiana','Pacers')
LAC = Team('LA Clippers','Clippers')
LAL = Team('LA Lakers','Lakers')
MEM = Team('Memphis','Grizzlies')
MIA = Team('Miami','Heat')
MIL = Team('Milwaukee','Bucks')
MIN = Team('Minnesota','Timberwolves')
NOP = Team('New Orleans','Pelicans')
NYK = Team('New York','Knicks')
OKC = Team('Okla City','Thunder')
ORL = Team('Orlando','Magic')
PHI = Team('Philadelphia','76ers')
PHX = Team('Phoenix','Suns')
POR = Team('Portland','Trail Blazers')
SAC = Team('Sacramento','Kings')
SAN = Team('San Antonio','Spurs')
TOR = Team('Toronto','Raptors')
UTA = Team('Utah','Jazz')
WAS = Team('Washington','Wizards')
Teams = [ATL,BOS,BKN,CHA,CHI,CLE,DAL,DEN,DET,GSW,HOU,IND,LAC,LAL,MEM,MIA,MIL,MIN,NOP,NYK,OKC,ORL,PHI,PHX,POR,SAC,SAN,TOR,UTA,WAS]

def SIMULATION(x,y):
    x_pts = 0
    y_pts = 0
    total_poss = (x.pace + y.pace)/2
    x_poss = []
    y_poss = []
    poss_toX = 0
    x_pace = 0
    y_pace = 0
    while poss_toX < total_poss:
        event = random.randint(1,100)
        if x.tno >= event:
            x_poss.append(0)
            poss_toX += 1
        else:
            x_poss.append(1)
            poss_toX += 1
    poss_toY = 0
    while poss_toY < total_poss:
        event = random.randint(1,100)
        if y.tno >= event:
            y_poss.append(0)
            poss_toY += 1
        else:
            y_poss.append(1)
            poss_toY += 1
    
    x_poss_total = sum(x_poss)
    y_poss_total = sum(y_poss)
    x_pace += x_poss_total
    y_pace += y_poss_total
    
    x_orebs = []
    y_orebs = []
    poss_orrX = 0
    while poss_orrX < x_poss_total:
        event = random.randint(1,1000)
        event = event/1000
        if y.drr < event:
            event = random.randint(1,1000)
            event = event/1000
            if x.orr > event:
                x_orebs.append(1)
                poss_orrX += 1
            else:
                x_orebs.append(0)
                poss_orrX += 1
        else:
            x_orebs.append(0)
            poss_orrX += 1
    poss_orrY = 0
    while poss_orrY < y_poss_total:
        event = random.randint(1,1000)
        event = event/1000
        if x.drr < event:
            event = random.randint(1,1000)
            event = event/1000
            if y.orr > event:
                y_orebs.append(1)
                poss_orrY += 1
            else:
                y_orebs.append(0)
                poss_orrY += 1
        else:
            y_orebs.append(0)
            poss_orrY += 1
            
    x_total_orebs = sum(x_orebs)
    y_total_orebs = sum(y_orebs)
    
    x_pace += x_total_orebs
    y_pace += y_total_orebs
        
    x_3_attempts = []
    y_3_attempts = []
    
    poss_3_x = 0
    while poss_3_x < x_pace:
        event = random.randint(1,1000)
        event = event/1000
        if x.threeRate > event:
            x_3_attempts.append(1)
            poss_3_x += 1
        else:
            x_3_attempts.append(0)
            poss_3_x += 1
            
    poss_3_y = 0
    while poss_3_y < y_pace:
        event = random.randint(1,1000)
        event = event/1000
        if y.threeRate > event:
            y_3_attempts.append(1)
            poss_3_y += 1
        else:
            y_3_attempts.append(0)
            poss_3_y += 1
            
        x_3_attempts_total = sum(x_3_attempts)
        y_3_attempts_total = sum(y_3_attempts)
    
    x_3_pts = []
    y_3_pts = []
    
    x_3_shots = 0
    while x_3_shots < x_3_attempts_total:
        event = random.randint(1,1000)
        event = event/1000
        if x.threePct > event:
            x_3_pts.append(3)
            x_3_shots += 1
        else:
            x_3_pts.append(0)
            x_3_shots += 1
            
    y_3_shots = 0
    while y_3_shots < y_3_attempts_total:
        event = random.randint(1,1000)
        event = event/1000
        if y.threePct > event:
            y_3_pts.append(3)
            y_3_shots += 1
        else:
            y_3_pts.append(0)
            y_3_shots += 1
            
    x_pts += sum(x_3_pts)
    y_pts += sum(y_3_pts)
    
    x_2_pts = []
    y_2_pts = []
    
    x_2_shots = 0
    while x_2_shots < x_pace - x_3_attempts_total:
        event = random.randint(1,1000)
        event = event/1000
        if x.twoPct > event:
            x_2_pts.append(2)
            x_2_shots += 1
        else:
            x_2_pts.append(0)
            x_2_shots += 1
    
    y_2_shots = 0
    while y_2_shots < y_pace - y_3_attempts_total:
        event = random.randint(1,1000)
        event = event/1000
        if y.twoPct > event:
            y_2_pts.append(2)
            y_2_shots += 1
        else:
            y_2_pts.append(0)
            y_2_shots += 1
            
    x_pts += sum(x_2_pts)
    y_pts += sum(y_2_pts)
    
    X_tno = len(x_poss)-sum(x_poss)
    Y_tno = len(y_poss)-sum(y_poss)
    X_orebs = sum(x_orebs)
    Y_orebs = sum(y_orebs)
    X_3_made = sum(x_3_pts)/3
    Y_3_made = sum(y_3_pts)/3
    X_3_pct = round( X_3_made/x_3_attempts_total,5)
    Y_3_pct = round(Y_3_made/y_3_attempts_total,5)
    X_2_shots = x_pace - x_3_attempts_total
    Y_2_shots = y_pace - y_3_attempts_total
    X_2_made = sum(x_2_pts)/2
    Y_2_made = sum(y_2_pts)/2
    X_2_pct = round(X_2_made/X_2_shots,5)
    Y_2_pct = round(Y_2_made/Y_2_shots,5)
    X_total_shots = x_3_attempts_total + X_2_shots
    Y_total_shots = y_3_attempts_total + Y_2_shots
    X_total_made = X_3_made + X_2_made
    Y_total_made = Y_3_made + Y_2_made
    X_total_pct = round(X_total_made/X_total_shots,5)
    Y_total_pct = round(Y_total_made/Y_total_shots,5)
    X_pts_per_shot = round(x_pts/X_total_shots,5)
    Y_pts_per_shot = round(y_pts/Y_total_shots,5)
    trace = go.Table(header={'values':['    ',x.mascot,y.mascot],
                            'fill_color':'lightgrey'
                            },
                    cells={'values':[['<b>Final Score','<b>Successful Possessions','<b>Turnovers','<b>Offensive Rebounds',
                                      '<b>Threes Attempted','<b>Threes Made','<b>Three Point Pct',
                                      '<b>Twos Attempted','<b>Twos Made','<b>Two Point Pct','<b>Total Shot Attempts',
                                     '<b>Total Shots Made','<b>Total Shooting Pct','<b>Pts Per Shot'],
                                     [x_pts,x_pace,X_tno,X_orebs,x_3_attempts_total,X_3_made,X_3_pct,X_2_shots,
                                      X_2_made,X_2_pct,X_total_shots,X_total_made,X_total_pct,X_pts_per_shot],
                                     [y_pts,y_pace,Y_tno,Y_orebs,y_3_attempts_total,Y_3_made,Y_3_pct,Y_2_shots,
                                     Y_2_made,Y_2_pct,Y_total_shots,Y_total_made,Y_total_pct,Y_pts_per_shot]]
                          })
    data = [trace]
    iplot({'data':data})


SIMULATION(PHI,MIA)

