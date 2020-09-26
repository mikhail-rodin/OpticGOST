fileName$ = $FILENAME()
fNameLen = SLEN(fileName$)
!remove extension
fName$ = $LEFTSTRING(fileName$, fNameLen - 4)
!now fName in filename without exptension

zmxPath$ = $PATHNAME()
! where zmx file is stored

FilePath$ = zmxPath$ + "\" + fName$ + "_config.txt"
OUTPUT FilePath$
PRINT "# this is a config file for the jsonexport macro"
PRINT "# modify it to have aberration data calculated for the specific rays you need"
PRINT "# aberrations are calculated for every combination of Px/Py/Hx/Hy"
PRINT "# use spaces as delimiters, use '# ' for comments"
PRINT "# ASCII encoding"
PRINT "Px_count: 3"
PRINT "Px: 0 0.7 1"
PRINT "Py_count: 4"
PRINT "Py: 0 0.5 0.7 1"

CONVERTFILEFORMAT FilePath$, 1 