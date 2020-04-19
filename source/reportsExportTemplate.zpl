!export raytrace
fileName$ = $FILENAME()
fNameLen = SLEN(fileName$)
!remove extension
fName$ = $LEFTSTRING(fileName$, fNameLen - 4)
zmxPath$ = $PATHNAME()
folder$ = zmxPath$ + "\"
msg$ = "Files saved to " + folder$
print msg$
configFolder$ = #configpath#  

fPath$ = folder$ + fName$ + "_raytraceAxial.TXT"
configPath$ = configFolder$ + "rtr_axial.CFG"
openanalysiswindow "rtr", configPath$
currentWindow = WINL()
savewindow currentWindow, fPath$
closewindow currentWindow

fPath$ = folder$ + fName$ + "_raytraceChief.TXT"
configPath$ = configFolder$ + "rtr_chief.CFG"
openanalysiswindow "rtr", configPath$
currentWindow = WINL()
savewindow currentWindow, fPath$
closewindow currentWindow

fPath$ = folder$ + fName$ + "_raytraceLower.TXT"
configPath$ = configFolder$ + "rtr_lower.CFG"
openanalysiswindow "rtr", configPath$
currentWindow = WINL()
savewindow currentWindow, fPath$
closewindow currentWindow

fPath$ = folder$ + fName$ + "_raytraceUpper.TXT"
configPath$ = configFolder$ + "rtr_upper.CFG"
openanalysiswindow "rtr", configPath$
currentWindow = WINL()
savewindow currentWindow, fPath$
closewindow currentWindow

fPath$ = folder$ + fName$ + "_FieldCurvDist.BMP"
configPath$ = configFolder$ + "fcd_wave1.CFG"
openanalysiswindow "fcd", configPath$
currentWindow = WINL()
EXPORTBMP currentWindow, fPath$, 500 #delay =500
closewindow currentWindow
fPath$ = folder$ + fName$ + "_FieldCurvDist.TXT"
GETTEXTFILE fPath$, fcd, configPath$, 1

fPath$ = folder$ + fName$ + "_ChromaticFocalShift.BMP"
!configPath$ = configFolder$ + "rtr_upper.CFG"
openanalysiswindow "cfs" #, configPath$
currentWindow = WINL()
EXPORTBMP currentWindow, fPath$, 500 #delay =500
closewindow currentWindow
fPath$ = folder$ + fName$ + "_ChromaticFocalShift.TXT"
GETTEXTFILE fPath$, cfs, configPath$. 1

fPath$ = folder$ + fName$ + "_Longitudinal.BMP"
!configPath$ = configFolder$ + "rtr_upper.CFG"
openanalysiswindow "lon" #, configPath$
currentWindow = WINL()
EXPORTBMP currentWindow, fPath$, 500 #delay =500
closewindow currentWindow
fPath$ = folder$ + fName$ + "_Longitudinal.TXT"
GETTEXTFILE fPath$, lon, configPath$, 1

fPath$ = folder$ + fName$ + "_Spherical.BMP"
configPath$ = configFolder$ + "ray_wave1field1.CFG"
openanalysiswindow "ray", configPath$
currentWindow = WINL()
EXPORTBMP currentWindow, fPath$, 500 #delay =500
closewindow currentWindow
fPath$ = folder$ + fName$ + "__Spherical.TXT"
GETTEXTFILE fPath$, ray, configPath$, 1

fPath$ = folder$ + fName$ + "_TR_Field2.BMP"
configPath$ = configFolder$ + "ray_wave1field2.CFG"
openanalysiswindow "ray", configPath$
currentWindow = WINL()
EXPORTBMP currentWindow, fPath$, 500 #delay =500
closewindow currentWindow
fPath$ = folder$ + fName$ + "_TR_Field2.TXT"
GETTEXTFILE fPath$, ray, configPath$, 1

fPath$ = folder$ + fName$ + "_TR_Field3.BMP"
configPath$ = configFolder$ + "ray_wave1field3.CFG"
openanalysiswindow "ray", configPath$
currentWindow = WINL()
EXPORTBMP currentWindow, fPath$, 500 #delay =500
closewindow currentWindow
fPath$ = folder$ + fName$ + "_TR_Field3.TXT"
GETTEXTFILE fPath$, ray, configPath$, 1
