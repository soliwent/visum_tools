#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:    ?????????Net???????????????????
#
# Author:      solin
#
# Created:     10/13/2018
# Copyright:   (c) solin 2018
# Licence:     <your licence>
#-------------------------------------------------------------------------------

# -*- coding: utf-8 -*-
import win32com.client as com
import sys,os
import time
Visum = com.Dispatch("Visum.Visum.150")
Visum.LoadVersion("D:/myPTV/Halle.ver")

t1= time.time()
Visum.Filters.InitAll() #??????
Visum.Filters.LinkFilter().AddCondition("OP_NONE",False,"TSYSSET","ContainsAll","CAR")  #??????????CAR???

avgSpeed = Visum.Net.AttValue(r"AvgActive:Links\v0PrT")

cntAll = Visum.Net.Links.Count
cntActive = Visum.Net.Links.CountActive
t2 = time.time()
print(cntAll)
print(cntActive)
print(avgSpeed)
print(str(t2-t1))
