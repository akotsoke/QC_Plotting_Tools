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
parser.add_argument('file', nargs='+')

args = parser.parse_args()


text_file = open("MyConfig.cfg", "w")
j=0

text_file.write("[BEGIN_CANVAS]\n")
text_file.write("    Canv_Axis_NDiv = '508,510'; #X,Y\n")
text_file.write("    Canv_Dim= '1000,1000'; #X,Y\n")
text_file.write("    Canv_DrawOpt = 'APE1';\n")
text_file.write("    Canv_Grid_XY = 'false,false';\n")
text_file.write("    Canv_Latex_Line = '0.62,0.86, #splitline{LS2}{Detector~Production}';\n")
text_file.write("    Canv_Latex_Line = '0.62,0.27, Gas~=~CO_{2}';\n")
text_file.write("    Canv_Latex_Line = '0.62,0.20, R_{Equiv}~=~5.000~M#Omega';\n")
text_file.write("    Canv_Legend_Dim_X = '0.20,0.60';\n")
text_file.write("    Canv_Legend_Dim_Y = '0.56,0.92';\n")
text_file.write("    Canv_Legend_Draw = 'true'; \n")
text_file.write("    Canv_Log_XY = 'false,false';\n")
text_file.write("    Canv_Logo_Pos = '0'; #0, 11, 22, 33\n")
text_file.write("    Canv_Logo_Prelim = 'true';\n")
text_file.write("    Canv_Margin_Top = '0.08';\n")
text_file.write("    Canv_Margin_Bot = '0.14';\n")
text_file.write("    Canv_Margin_Lf = '0.16';\n")
text_file.write("    Canv_Margin_Rt = '0.06';\n")
text_file.write("    Canv_Name = 'Ls2_QC4_Imon_vs_ApplVoltage_AllDet';\n")
text_file.write("    Canv_Plot_Type = 'TGraphErrors';\n")
text_file.write("    Canv_Range_X = '0,1000'; #X1, X2\n")
text_file.write("    Canv_Range_Y = '0,7'; #Y1, Y2\n")
text_file.write("    Canv_Title_Offset_X = '1.0';\n")
text_file.write("    Canv_Title_Offset_Y = '0.8';\n")
text_file.write("    Canv_Title_X = 'Divider Current #left(#muA#right)';\n")
text_file.write("    Canv_Title_Y = 'Applied Voltage #left(kV#right)';\n")
for f in args.file:
	
	workbook = xlrd.open_workbook(str(f))
	ws = workbook.sheet_by_index(0)

	j=j+1
	Marker_Style = int(20+j)
	text_file.write("    [BEGIN_PLOT]\n" )
	text_file.write("        Plot_Color = 'kRed+"+str(j)+  "';\n" )
	text_file.write("        Plot_LegEntry = '';\n" )
	text_file.write("        Plot_Line_Size = '1';\n" )
	text_file.write("        Plot_Line_Style = '1';\n" )
	text_file.write("        Plot_Marker_Size = '1';\n" )
	text_file.write("        Plot_Marker_Style = '"+str(Marker_Style)+"';\n" )
	text_file.write("        Plot_Name = ''; \n" )
	text_file.write("        [BEGIN_DATA]\n" )
	text_file.write("          VAR_INDEP,VAR_DEP\n" )

	for i in range(3,37):
		Vmon= float(ws.cell_value(i,1)/1000)
		Imon= ws.cell_value(i,3)
		text_file.write(str(Imon)+","+str(Vmon)+"\n")
		pass
         
	text_file.write("        [END_DATA]\n" )
	text_file.write("     [END_PLOT]\n" )

text_file.write("[END_CANVAS]\n" )
text_file.close()