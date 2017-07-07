
import pandas as pd
import xlrd
from array import array
import numpy as np
from ROOT import TGraph
from ROOT import TMultiGraph
import ROOT as r

import argparse
from optparse import OptionParser

parser = argparse.ArgumentParser()
parser.add_argument('filename', nargs='+')



args = parser.parse_args()
#print args.filename
#filename = raw_input("> Enter the Input Filename: ")
#Det = raw_input("> Enter the Detector Name: ")

for f in args.filename:
	workbook = xlrd.open_workbook(str(f))
	ws = workbook.sheet_by_index(0)

	

	
	time= [ws.cell_value(i,1) for i in range(1,61)]
	pres= [ws.cell_value(i,2) for i in range(1,61)]
	temp= [ws.cell_value(i,3) for i in range(1,61)]
	atmP= [ws.cell_value(i,4) for i in range(1,61)]
	
	 
	
	

	time = np.array(time)
	pres=np.array(pres)
	temp=np.array(temp)
	atmP=np.array(atmP)
	

	canv = r.TCanvas("canv","canv",700,600)
	canv.cd()



	pad1 = r.TPad("pad1","",0,0,1,1)
	pad2= r.TPad("pad2","",0,0,1,1)
	pad2.SetFillStyle(4000)
	pad2.SetFrameFillStyle(0)
	pad1.Draw()
	pad1.cd()

	#============Plot Inner Pressure vs Time ===============
	gr= TGraph(60,time,pres)
	gr.Draw("AP")

	# For X axis
	gr.GetXaxis().SetNdivisions(505)
	gr.GetXaxis().SetTitle("Time (s)")
	gr.GetXaxis().SetRangeUser(0,3600)
	gr.GetXaxis().SetTitleSize(0.06)
	gr.GetXaxis().SetLabelSize(0.05)
	gr.GetXaxis().SetLabelOffset(0.007)
	gr.GetXaxis().SetLabelFont(42)
	gr.GetXaxis().SetLabelColor(1)
	gr.GetXaxis().SetTitleOffset(0.9)
	gr.GetXaxis().SetAxisColor(1)
	r.TStyle().SetStripDecimals(r.kTRUE)
	gr.GetXaxis().SetTickLength(0.03)
	r.TStyle().SetPadTickX(1)


	# For Y axis
	gr.GetYaxis().SetTitle("#splitline{Pressure (mbar)}{ Temperature (C) }")
	gr.GetYaxis().SetNdivisions(508)
	gr.GetYaxis().SetTitleOffset(1.1)
	gr.GetYaxis().SetRangeUser(0,35)
	gr.GetYaxis().SetTitleSize(0.06)
	gr.GetYaxis().SetLabelSize(0.05)
	gr.GetYaxis().SetLabelOffset(0.007)
	gr.GetYaxis().SetLabelFont(42)
	gr.GetYaxis().SetLabelColor(1)
	gr.GetYaxis().SetTitleOffset(1.1)
	gr.GetYaxis().SetAxisColor(1)

	r.TStyle().SetStripDecimals(r.kTRUE)
	gr.GetXaxis().SetTickLength(0.03)
	r.TStyle().SetPadTickX(1)


	gr.SetTitle("")
	gr.SetMarkerStyle(20)
	gr.SetMarkerSize(0.8)
	gr.SetMarkerColor(r.kBlue)
	gr.SetLineStyle(1)




	#============Plot Temperature vs Time ===============
	gr_T= TGraph(60,time,temp)
	gr_T.SetMarkerStyle(20)
	gr_T.SetMarkerSize(0.8)
	gr_T.SetMarkerColor(r.kOrange)
	gr_T.SetLineStyle(1)
	gr_T.Draw("PSAME")


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
	gr_Atm.GetYaxis().SetTitle("Atm Pressure (mbar)")
	gr_Atm.GetYaxis().SetTitleSize(0.05)
	gr_Atm.GetYaxis().SetTitleOffset(1.1)

	# For X axis
	gr_Atm.GetXaxis().SetLabelOffset(999)
	gr_Atm.GetXaxis().SetLabelSize(0)
	gr_Atm.GetXaxis().SetTickLength(0)
	r.TStyle().SetPadTickX(1)


	# For Y axis

	gr_Atm.GetYaxis().SetNdivisions(508)
	gr_Atm.GetYaxis().SetTitleOffset(1.2)
	gr_Atm.GetYaxis().SetTitleSize(0.06)
	gr_Atm.GetYaxis().SetLabelSize(0.05)
	gr_Atm.GetYaxis().SetLabelOffset(0.007)
	gr_Atm.GetYaxis().SetLabelFont(42)
	gr_Atm.GetYaxis().SetLabelColor(1)
	gr_Atm.GetYaxis().SetTitleOffset(1.1)
	gr_Atm.GetYaxis().SetAxisColor(1)

	r.TStyle().SetStripDecimals(r.kTRUE)
	gr.GetXaxis().SetTickLength(0.03)
	r.TStyle().SetPadTickX(1)




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

	# For the Legend
	pad2.cd()
	leg= r.TLegend()
	leg.SetX1NDC(0.55)
	leg.SetX2NDC(0.80)
	leg.SetY1NDC(0.17)
	leg.SetY2NDC(0.42)
	leg.SetBorderSize(0)
	leg.SetFillStyle(0)
	leg.SetTextSize(0.04)
	leg.AddEntry(gr,"Pressure","lep")
	leg.AddEntry(gr_T,"Temperature","lep")
	leg.AddEntry(gr_Atm,"Atm Pressure","lep")
	leg.Draw()

	# For the Pad
	r.TStyle().SetPadBorderMode(0);
	r.TStyle().SetPadBorderSize(1);
	r.TStyle().SetPadColor(r.kWhite);
	r.TStyle().SetPadGridX(r.false);
	r.TStyle().SetPadGridY(r.false);
	r.TStyle().SetGridColor(0);
	r.TStyle().SetGridStyle(3);
	r.TStyle().SetGridWidth(1);

	# For the frame:

	r.TStyle().SetFrameBorderMode(0);
	r.TStyle().SetFrameBorderSize(1);
	r.TStyle().SetFrameFillColor(0);
	r.TStyle().SetFrameFillStyle(0);
	r.TStyle().SetFrameLineColor(1);
	r.TStyle().SetFrameLineStyle(1);
	r.TStyle().SetFrameLineWidth(1);

	#========= Latex lines : logo etc =============

	logo= r.TLatex(0.6,0.8,"")
	logo.SetNDC()
	logo.DrawLatex(0.17,0.93,"#scale[1.2]{CMS} #font[52]{Preliminary}")
	logo.SetTextFont(4)
	pad2.Update()
	#Tline = r.TLatex(0.62,0.78,"")
	#Tline.SetNDC()
	#Tline.SetTextSize(0.04)
	#Tline.DrawLatex(0.18,0.87, "#font[42]{" + str(Det)+ "}")
	
	#Tline.SetTextAlign(12)
	#Tline.Draw()
	
	#==========================================================
	b="xlsm"
	
	a = str(f)
	a= a.replace(b,"png")
	#print a
	#aprint str(f)

	canv.SaveAs(str(a))

	pass





