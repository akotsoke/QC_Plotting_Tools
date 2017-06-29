
import pandas as pd
import xlrd
from array import array
import numpy as np
from ROOT import TGraph
from ROOT import TMultiGraph
import ROOT as r

#file = 'QC3_GE11-X-S-CERN-0001_20170503_test6.xlsm'
#data= pd.ExcelFile(file)

#df1= data.parse(0,skiprows=1)
#print(df1.as_matrix(columns=[1,4]))

filename = raw_input("> Enter the Input Filename: ")
Det = raw_input("> Enter the Detector Name: ")


workbook = xlrd.open_workbook(str(filename))
ws = workbook.sheet_by_index(0)

print ws.cell(0, 0).value



#time = []

time= [ws.cell_value(i,1) for i in range(1,61)]
pres= [ws.cell_value(i,2) for i in range(1,61)]
temp= [ws.cell_value(i,3) for i in range(1,61)]
atmP= [ws.cell_value(i,4) for i in range(1,61)]
#value= str(float(value))
#print value
#time.append(value)


time = np.array(time)
pres=np.array(pres)
temp=np.array(temp)
atmP=np.array(atmP)


canv = r.TCanvas("canv","canv",700,600)
canv.cd()

#mg = TMultiGraph("mg","")

pad1 = r.TPad("pad1","",0,0,1,1)
pad2= r.TPad("pad2","",0,0,1,1)
pad2.SetFillStyle(4000)
pad2.SetFrameFillStyle(0)
pad1.Draw()
pad1.cd()
#============Plot Inner Pressure vs Time ===============
gr= TGraph(60,time,pres)
gr.Draw("AP")
gr.GetXaxis().SetNdivisions(510)
gr.GetYaxis().SetNdivisions(510)
gr.GetXaxis().SetTitle("Time (s)")
gr.GetYaxis().SetTitle("#splitline{Pressure (mbar)}{ Temperature (C) }")
gr.GetXaxis().SetTitleOffset(1)
gr.GetXaxis().SetTitleSize(0.05)
gr.GetYaxis().SetTitleSize(0.05)
gr.GetYaxis().SetTitleOffset(1.1)
gr.GetYaxis().SetRangeUser(0,35)
gr.GetXaxis().SetRangeUser(0,3600)
gr.SetTitle("")
gr.SetMarkerStyle(20)
gr.SetMarkerSize(0.8)
gr.SetMarkerColor(r.kBlue)
gr.SetLineStyle(1)
#mg.Add(gr,"P")


#============Plot Temperature vs Time ===============
gr_T= TGraph(60,time,temp)
gr_T.SetMarkerStyle(20)
gr_T.SetMarkerSize(0.8)
gr_T.SetMarkerColor(r.kOrange)
gr_T.SetLineStyle(1)
#mg.Add(gr_T,"P")
gr_T.Draw("PSAME")
pad1

#============Plot Atm Pressure vs Time ===============
pad2.Draw()
pad2.cd()
gr_Atm= TGraph(60,time,atmP)
gr_Atm.SetMarkerStyle(20)
gr_Atm.SetMarkerSize(0.8)
gr_Atm.SetMarkerColor(r.kGray)
gr_Atm.SetLineStyle(1)
gr_Atm.SetTitle("")
gr_Atm.Draw("PAY+")
gr_Atm.GetYaxis().SetRangeUser(900,1000)
gr_Atm.GetXaxis().SetRangeUser(0,3600)
gr_Atm.GetYaxis().SetTitle("Atm Pressure (mbar)")
gr_Atm.GetYaxis().SetTitleSize(0.05)
gr_Atm.GetYaxis().SetTitleOffset(1.1)

#mg.Add(gr_Atm,"LPY")
#mg.Draw("AP")




#===================Canvas styling==========================
pad1.cd()
r.gPad.SetLeftMargin(0.16)
r.gPad.SetRightMargin(0.16)
r.gPad.SetTopMargin(0.08)
r.gPad.SetBottomMargin(0.14)

pad2.cd()
r.gPad.SetLeftMargin(0.16)
r.gPad.SetRightMargin(0.16)
r.gPad.SetTopMargin(0.08)
r.gPad.SetBottomMargin(0.14)

pad2.cd()
leg= r.TLegend()
leg.SetX1NDC(0.55)
leg.SetX2NDC(0.80)
leg.SetY1NDC(0.17)
leg.SetY2NDC(0.42)
leg.AddEntry(gr,"Pressure","lep")
leg.AddEntry(gr_T,"Temperature","lep")
leg.AddEntry(gr_Atm,"Atm Pressure","lep")
leg.Draw()


logo= r.TLatex(0.6,0.8,"")
logo.SetNDC()
logo.DrawLatex(0.17,0.93,"#scale[1.2]{CMS} #font[52]{Preliminary}")
logo.SetTextFont(4)
pad2.Update()
Tline = r.TLatex(0.61,0.78,"")
Tline.SetNDC()
Tline.SetTextSize(0.04)
Tline.DrawLatex(0.17,0.87, "#font[42]{" + str(Det)+ "}")
#
#Tline.SetTextAlign(12)
Tline.Draw()

#==========================================================
OutName = "QC3_"+ str(Det)+ ".png"
canv.SaveAs(OutName)


