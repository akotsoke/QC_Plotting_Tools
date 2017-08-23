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
#Read_File = []
Detector = []
x_err= np.zeros(34)

args = parser.parse_args()

#Read_File.append("GE11-X-L-CERN-0002_QC4_20170607-Test2.xlsx")
#Read_File.append("GE11-X-S-CERN-0001_QC4_20170524.xlsx")
#Read_File.append("GE11-X-S-CERN-0002_QC4_20170609.xlsx")
#Read_File.append("GE11-X-S-CERN-0003_QC4_20170613_Test_1.xlsx")

def Color(j):
	if j % 7 == 0:
		ret_int=r.kRed
	elif j % 7 == 1:
		ret_int=r.kRed+1
	elif j % 7 == 2:
		ret_int=r.kRed+2
	elif j % 7 == 3:
		ret_int=r.kRed+3
	elif j % 7 == 4:
		ret_int=r.kBlue
	elif j % 7 == 5:
		ret_int=r.kBlue+1
	elif j % 7 == 6:
		ret_int=r.kBlue+2
	

	return ret_int


mg = TMultiGraph("mg","")





leg= r.TLegend()
i=0

for f in args.filename:
	
	workbook = xlrd.open_workbook(str(f))
	ws = workbook.sheet_by_index(0)
	Gcharacter=int(f.index('G'))
	legEntry=f[int(Gcharacter):int(Gcharacter+18)]

	SS = [ws.cell_value(j,7) for j in range(3,37)]
	Imon = [ws.cell_value(j,3) for j in range(3,37)]
	SS_err= [ws.cell_value(j,8) for j in range(3,37)]

	SS = np.array(SS)
	Imon = np.array(Imon)
	SS_err = np.array(SS_err)

	
	gr = r.TGraphErrors(34,Imon,SS,x_err,SS_err)

	
	gr.Draw("AP")
	gr.SetTitle("")
	gr.SetMarkerStyle(20+i)
	gr.SetMarkerSize(0.8)
	gr.SetMarkerColor(Color(i))
	gr.SetLineStyle(1)
	gr.SetLineColor(Color(i))
    
	
	leg.AddEntry(gr,str(legEntry),"lep")



	mg.Add(gr,"P")

	i=i+1

	pass




#========Canvas Styling ===============

canv = r.TCanvas("canv","canv",600,600)
#r.TStyle.SetCanvasBorderMode(0)
r.TStyle().SetCanvasColor(r.kWhite)
r.TStyle().SetCanvasDefX(0)
r.TStyle().SetCanvasDefY(0)
canv.SetLeftMargin(0.16)
canv.SetRightMargin(0.06)
canv.SetTopMargin(0.08)
canv.SetBottomMargin(0.13)
canv.cd()


mg.Draw("AP SAME")
leg.Draw()




# For the legend
leg.SetBorderSize(0)
leg.SetFillStyle(0)
leg.SetX1NDC(0.18)
leg.SetX2NDC(0.60)
leg.SetY1NDC(0.55)
leg.SetY2NDC(0.90)
leg.SetTextSize(0.03)



# X Axis Style
mg.GetXaxis().SetTitle("Current Drawn (#mu A)")
mg.GetXaxis().SetTitleSize(0.06)
mg.GetXaxis().SetLabelSize(0.05)
mg.GetXaxis().SetNdivisions(508)
mg.GetXaxis().SetLabelOffset(0.007)
mg.GetXaxis().SetLabelFont(42)
mg.GetXaxis().SetLabelColor(1)
mg.GetXaxis().SetTitleOffset(0.9)
mg.GetXaxis().SetAxisColor(1)
r.TStyle().SetStripDecimals(r.kTRUE)
mg.GetXaxis().SetTickLength(0.03)
r.TStyle().SetPadTickX(1)


# Y Axis Style
mg.GetYaxis().SetTitle("Spurious Signal R_{SS} (Hz)")
mg.GetYaxis().SetRangeUser(0,50)
mg.GetYaxis().SetLabelSize(0.05)
mg.GetYaxis().SetLabelOffset(0.007)
mg.GetYaxis().SetLabelFont(42)
mg.GetYaxis().SetLabelColor(1)
mg.GetYaxis().SetTitleSize(0.06)
mg.GetYaxis().SetTitleOffset(1)
mg.GetYaxis().SetNdivisions(510)
r.TStyle().SetPadTickY(1)
mg.GetYaxis().SetAxisColor(1)
mg.GetYaxis().SetTickLength(0.03)
r.TStyle().SetPadTickY(1)

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

Tline = r.TLatex(0.61,0.78,"")
Tline.SetNDC()
Tline.SetTextSize(0.04)
Tline.DrawLatex(0.61,0.85, "#font[42]{Gas = CO_{2}}")

Tline1 = r.TLatex(0.61,0.78,"")
Tline1.SetNDC()
Tline1.SetTextSize(0.04)
Tline1.DrawLatex(0.61,0.80, "#font[42]{ReadOut = GEM3B}")

#Tline2 = r.TLatex(0.61,0.78,"")
#Tline2.SetNDC()
#Tline2.SetTextSize(0.04)
#Tline2.DrawLatex(0.18,0.87, "#font[42]{#splitline{LS2}{ Detector Production}}")







canv.SaveAs("QC4_SS_vs_Imon_Comparison.png")
