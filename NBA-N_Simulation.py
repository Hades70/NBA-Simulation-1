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
import re

book = xw.Book('data.csv')
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
Teams = [ATL,BOS,BKN,CHA,CHI,CLE,DAL,DEN,DET,GSW,HOU,IND,LAC,LAL,MEM,MIA,MIL,MIN,NOP,NYK,OKC,ORL,PHI,PHX,POR,SAC,
         SAN,TOR,UTA,WAS]


def N_SIMULATION_MC(x,y,n,xBook):
    Total_games = 0
    Final_score_x = []
    Final_score_y = []
    Pace_x = []
    Pace_y = []
    Successful_possesions_x = []
    Successful_possesions_y = []
    Turnovers_x = []
    Turnovers_y = []
    O_rebs_x = []
    O_rebs_y = []
    Threes_attempted_x = []
    Threes_attempted_y = []
    Threes_made_x = []
    Threes_made_y = []
    Three_point_pct_x = []
    Three_point_pct_y = []
    Twos_attempted_x = []
    Twos_attempted_y = []
    Twos_made_x = []
    Twos_made_y = []
    Two_point_pct_x = []
    Two_point_pct_y = []
    Total_shot_attempts_x = []
    Total_shot_attempts_y = []
    Total_shots_made_x = []
    Total_shots_made_y = []
    Total_shooting_pct_x = []
    Total_shooting_pct_y = []
    Pts_per_shot_x = []
    Pts_per_shot_y = []
    while Total_games < n:
        x_pts = 0
        y_pts = 0
        total_poss = int((x.pace + y.pace)/2)
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
        
        Final_score_x.append(x_pts)
        Final_score_y.append(y_pts)
        Pace_x.append((x.pace+y.pace)/2)
        Pace_y.append((x.pace+y.pace)/2)
        Successful_possesions_x.append(x_pace-X_tno)
        Successful_possesions_y.append(y_pace-Y_tno)
        Turnovers_x.append(X_tno)
        Turnovers_y.append(Y_tno)
        O_rebs_x.append(X_orebs)
        O_rebs_y.append(Y_orebs)
        Threes_attempted_x.append(x_3_attempts_total)
        Threes_attempted_y.append(y_3_attempts_total)
        Threes_made_x.append(X_3_made)
        Threes_made_y.append(Y_3_made)
        Three_point_pct_x.append(X_3_pct)
        Three_point_pct_y.append(Y_3_pct)
        Twos_attempted_x.append(X_2_shots)
        Twos_attempted_y.append(Y_2_shots)
        Twos_made_x.append(X_2_made)
        Twos_made_y.append(Y_2_made)
        Two_point_pct_x.append(X_2_pct)
        Two_point_pct_y.append(Y_2_pct)
        Total_shot_attempts_x.append(x_3_attempts_total+X_2_shots)
        Total_shot_attempts_y.append(y_3_attempts_total+Y_2_shots)
        Total_shots_made_x.append(X_total_made)
        Total_shots_made_y.append(Y_total_made)
        Total_shooting_pct_x.append(X_total_pct)
        Total_shooting_pct_y.append(Y_total_pct)
        Pts_per_shot_x.append(X_pts_per_shot)
        Pts_per_shot_y.append(Y_pts_per_shot)
    
        
        Total_games += 1
    X_name = str(x.mascot)
    Y_name = str(y.mascot)
    x_tick = []
    K = 0
    for i in range(1,n+1,1):
        K +=1
        x_tick.append(K)
        
    xBook = xw.Book()
    sht = xBook.sheets('Sheet1')
    zx = sht.range
    
    zx('a1').value = 'Game_ID'
    zx('b1').value = X_name + '_Points'
    zx('c1').value = Y_name + '_Points'
    zx('d1').value = X_name + '_Succesful_Possessions'
    zx('e1').value = Y_name + '_Succesful_Possessions'
    zx('f1').value = X_name + '_Turnovers'
    zx('g1').value = Y_name + '_Turnovers'
    zx('h1').value = X_name + '_O_Rebs'
    zx('i1').value = Y_name + '_O_Rebs'
    zx('j1').value = X_name + '_Threes_Attempted'
    zx('k1').value = Y_name + '_Threes_Attempted'
    zx('l1').value = X_name + '_Threes_Made'
    zx('m1').value = Y_name + '_Threes_Made'
    zx('n1').value = X_name + '_Three_Point%'
    zx('o1').value = Y_name + '_Three_Point%'
    zx('p1').value = X_name + '_Twos_Attempted'
    zx('q1').value = Y_name + '_Twos_Attempted'
    zx('r1').value = X_name + '_Twos_Made'
    zx('s1').value = Y_name + '_Twos_Made'
    zx('t1').value = X_name + '_Two_Point%'
    zx('u1').value = Y_name + '_Two_Point%'
    zx('v1').value = X_name + '_Total_Shots_Attempted'
    zx('w1').value = Y_name + '_Total_Shots_Attempted'
    zx('x1').value = X_name + '_Total_Shots_Made'
    zx('y1').value = Y_name + '_Total_Shots_Made'
    zx('z1').value = X_name + '_Total_Shooting%'
    zx('aa1').value = Y_name + '_Total_Shooting%'
    zx('ab1').value = X_name + '_Pts_Per_shot'
    zx('ac1').value = Y_name + '_Pts_Per_shot'
    
    zx('a2').options(transpose=True).value = x_tick
    zx('b2').options(transpose=True).value = Final_score_x
    zx('c2').options(transpose=True).value = Final_score_y
    zx('d2').options(transpose=True).value = Successful_possesions_x
    zx('e2').options(transpose=True).value = Successful_possesions_y
    zx('f2').options(transpose=True).value = Turnovers_x
    zx('g2').options(transpose=True).value = Turnovers_y
    zx('h2').options(transpose=True).value = O_rebs_x
    zx('i2').options(transpose=True).value = O_rebs_y
    zx('j2').options(transpose=True).value = Threes_attempted_x
    zx('k2').options(transpose=True).value = Threes_attempted_y
    zx('l2').options(transpose=True).value = Threes_made_x
    zx('m2').options(transpose=True).value = Threes_made_y
    zx('n2').options(transpose=True).value = Three_point_pct_x
    zx('o2').options(transpose=True).value = Three_point_pct_y
    zx('p2').options(transpose=True).value = Twos_attempted_x
    zx('q2').options(transpose=True).value = Twos_attempted_y
    zx('r2').options(transpose=True).value = Twos_made_x
    zx('s2').options(transpose=True).value = Twos_made_y
    zx('t2').options(transpose=True).value = Two_point_pct_x
    zx('u2').options(transpose=True).value = Two_point_pct_y
    zx('v2').options(transpose=True).value = Total_shot_attempts_x
    zx('w2').options(transpose=True).value = Total_shot_attempts_y
    zx('x2').options(transpose=True).value = Total_shots_made_x
    zx('y2').options(transpose=True).value = Total_shots_made_y
    zx('z2').options(transpose=True).value = Total_shooting_pct_x
    zx('aa2').options(transpose=True).value = Total_shooting_pct_y
    zx('ab2').options(transpose=True).value = Pts_per_shot_x
    zx('ac2').options(transpose=True).value = Pts_per_shot_y
    
    listX = []
    listY = []
    str_ticks = [str(i) for i in x_tick]
    dict_scores_x = dict(zip(str_ticks,Final_score_x))
    dict_scores_y = dict(zip(str_ticks,Final_score_y))
    
    for key,value in dict_scores_x.items():
        for akey,avalue in dict_scores_y.items():
            if key == akey:
                if value > avalue:
                    listX.append(1)
                    listY.append(0)
                if value < avalue:
                    listX.append(0)
                    listY.append(1)
                else:
                    pass
    
    Win_pct_x = sum(listX)/len(listX)
    Win_pct_y = sum(listY)/len(listY)
    
    px = [Final_score_x,Turnovers_x,O_rebs_x,Threes_attempted_x,Threes_made_x,Three_point_pct_x,
          Twos_attempted_x,Twos_made_x,Two_point_pct_x,Total_shot_attempts_x,Total_shots_made_x,Total_shooting_pct_x,
          Pts_per_shot_x]
    py = [Final_score_y,Turnovers_y,O_rebs_y,Threes_attempted_y,Threes_made_y,Three_point_pct_y,
          Twos_attempted_y,Twos_made_y,Two_point_pct_y,Total_shot_attempts_y,Total_shots_made_y,Total_shooting_pct_y,
          Pts_per_shot_y]

    Px = []
    Py = []
    
    for i in px:
        Px.append(round(np.mean(i),5))
    for j in py:
        Py.append(round(np.mean(j),5))
    Px.append(round(Win_pct_x,5))
    Py.append(round(Win_pct_y,5))
    
    titles = ['<b>Final Score','<b>Turnovers','<b>Offensive Rebounds','<b>Threes Attempted','<b>Threes Made',
              '<b>Three Point Pct','<b>Twos Attempted','<b>Twos Made','<b>Two Point Pct','<b>Total Shot Attempts',
              '<b>Total Shots Made','<b>Total Shooting Pct','<b>Pts Per Shot']
    
    x_str = str(x_tick)
    p = 0
    while p < len(px)-1:
        for i in range(1,len(px)+1,1):
            if re.search('Made',titles[p]) is None and p != 1 and p != 2:
                trace = go.Scatter(x=x_tick,
                                  y=px[p],
                                  mode='markers',
                                  name=x.mascot,
                                   marker={'color':'orange'}
                                  )
                trace1 = go.Scatter(x=x_tick,
                                   y=py[p],
                                   mode='markers',
                                   name=y.mascot,
                                    marker={'color':'blue'}
                                   )
                layout = {'title':titles[p]}
                data = [trace,trace1]
                iplot({'data':data,'layout':layout})
                
                traceA = go.Histogram(x=px[p],
                                     name=x.mascot,
                                      marker={'color':'orange'}
                                     )
                traceB = go.Histogram(x=py[p],
                                     name=y.mascot,
                                      marker={'color':'blue'}
                                     )
                dataAB = [traceA,traceB]
                title = '<b>Distribution of '+titles[p]
                layoutA = {'title':title}
                iplot({'data':dataAB,'layout':layoutA})
                p += 1
            else:
                trace = go.Scatter(x=x_tick,
                                  y=px[p],
                                  mode='lines',
                                  name=x.mascot,
                                   marker={'color':'orange'}
                                  )
                trace1 = go.Scatter(x=x_tick,
                                   y=py[p],
                                   mode='lines',
                                   name=y.mascot,
                                    marker={'color':'blue'}
                                   )
                layout = {'title':titles[p]}
                data = [trace,trace1]
                iplot({'data':data,'layout':layout})
                
                traceA = go.Histogram(x=px[p],
                                     name=x.mascot,
                                      marker={'color':'orange'}
                                     )
                traceB = go.Histogram(x=py[p],
                                     name=y.mascot,
                                      marker={'color':'blue'}
                                     )
                dataAB = [traceA,traceB]
                title = '<b>Distribution of ' + titles[p]
                layoutA = {'title':title}
                iplot({'data':dataAB,'layout':layoutA})
                p += 1

    
    trace = go.Table(header={'values':['    ',x.mascot,y.mascot],
                            'fill_color':'lightgrey'
                            },
                     cells={'values':[
                         ['<b>Final Score','<b>Turnovers','<b>Offensive Rebounds','<b>Threes Attempted',
                          '<b>Threes Made','<b>Three Point Pct','<b>Twos Attempted','<b>Twos Made','<b>Two Point Pct',
                          '<b>Total Shot Attempts','<b>Total Shots Made','<b>Total Shooting Pct','<b>Pts Per Shot',
                          '<b>Percent Games Won'],
                         [i for i in Px],
                         [j for j in Py]
                                     ]}
                    )
    layout = {'title':'<b>Average Outcome'} 
    data = [trace]
    iplot({'data':data,'layout':layout})


N = 1000
B = None
N_SIMULATION_MC(LAL,MIL,N,B)

