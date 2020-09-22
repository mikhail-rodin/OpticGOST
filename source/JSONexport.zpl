! JSONexport
!
! This macro exports lens data (r,n,d, glass) and minimal
! aberration analysis data in a widespread easy-to-parse JSON format.
! Data can be imported into Excel to create GOST-compliant data sheets
! using the OpticGOST add-in. 
!
! The JSON file is saved alongside the zmx file. 
! A config.txt file has to be generated with a JSONconfig macro
! before exporting to JSON. Coordinates of rays traced to calculate 
! aberrations are specified in config.txt. Thus you have the flexibility
! to find chief ray aberrations for unvignetted field, for example. 
!
! https://github.com/mikhail-rodin/OpticGOST

fileName$ = $FILENAME()
fNameLen = SLEN(fileName$)
!remove extension
fName$ = $LEFTSTRING(fileName$, fNameLen - 4)
!now fName in filename without exptension

zmxPath$ = $PATHNAME()
! where zmx file is stored

jsonFilePath$ = zmxPath$ + "\" + fName$ + "_lensdata.json"

! read config file
DECLARE Hx, double, 1, 50
DECLARE Hy, double, 1, 50
DECLARE Px, double, 1, 50
DECLARE Py, double, 1, 50
Py(1) = 0 
Py(2) = 0.5
Py(3) = 0.707
Py(4) = 1
Px(1) = 0
Hx(1) = 0
Hy(1) = 1
Hx_count = 1
Hy_count = 1
Px_count = 1
Py_count = 4
configFilePath$ = zmxPath$ + "\" + fName$ + "_config.txt"
OPEN configFilePath$
LABEL 10
IF EOFF() THEN GOTO 20
READSTRING confLine$
first$ = $GETSTRING(confLine$,1)
IF first$ $== "#" THEN GOTO 10
!IF first$ $== "Hx_count:" 
    !Hx_count = SVAL($GETSTRING(confLine$,2))
    !GOTO 10
ELSE
    IF first$ $== "Hy_count:" 
        Hy_count = SVAL($GETSTRING(confLine$, 2))
        GOTO 10
    ELSE
        IF first$ $== "Px_count:" 
            Px_count = SVAL($GETSTRING(confLine$,2))
            GOTO 10
        ELSE
            IF first$ $== "Py_count:" 
                Py_count = SVAL($GETSTRING(confLine$,2))
                GOTO 10
            ELSE
                IF first$ $== "Hx" 
                    FOR i, 1, Hx_count, 1
                    Hx(i) = SVAL($GETSTRING(confLine$,i+1))
                    NEXT
                    GOTO 10
                ELSE
                   IF first$ $== "Hy" 
                       FOR i, 1, Hy_count, 1
                       Hy(i) = SVAL($GETSTRING(confLine$,i+1))
                       NEXT
                       GOTO 10
                    ELSE
                        IF first$ $== "Px" 
                            FOR i, 1, Px_count, 1
                            Px(i) = SVAL($GETSTRING(confLine$,i+1))
                            NEXT
                            GOTO 10
                        ELSE
                            IF first$ $== "Py" 
                                FOR i, 1, Py_count, 1
                                Py(i) = SVAL($GETSTRING(confLine$,i+1))
                                NEXT
                                GOTO 10
                            ENDIF
                        ENDIF
                    ENDIF
                ENDIF
            ENDIF
        ENDIF
    ENDIF
ENDIF
LABEL 20
CLOSE 


OUTPUT jsonFilePath$

!variables
!   str$  -  (string) temp for writing to file
!   val   -  (float) temp  
!   id    -  (int) operand 
!   noreturn  - (null) for calling functions that don't return anything useful (SYPR() etc)

!number of wavelengths
waveCount = NWAV()
primaryWave = PWAV()

afocal_im_space = SYPR(18)
! 0 = false, 1 = true
id = OCOD("EFFL")
! EFFL(void, wave)
IF (afocal_im_space == 0) THEN effl = OPEV(id, 0, primaryWave, 0, 0, 0, 0)

!number of fields
fieldCount = NFLD()
fieldTypeID = SYPR(100)
! 0 for angle
! 1 for obj height
! 2 for paraxial image height
! 3 for real image height
maxField = MAXF()

apertureType = SYPR(10)
! 0 for entrance pupil diameter
! 1 for image space F/#
! 2 for object space num aperture NA
! 3 for float by stop size
! 4 for paraxial working F/#
! 5 for object cone angle in degrees
apertureValue = SYPR(11)
id = OCOD("ENPD")
! ENPD(void)
exitPupilDiam = OPEV(id, 0, 0, 0, 0, 0, 0)
IF (apertureType == 0) 
    entrPupilDiam = apertureValue
ELSE
    IF (apertureType == 1)
        entrPupilDiam = effl/apertureValue
    ELSE
        id = OCOD("ENPD")
        ! EPDI(void)
        entrPupilDiam = OPEV(id, 0,0,0,0,0,0)
    ENDIF
ENDIF
id = OCOD("ENPP")
entrPupilPos = OPEV(id, 0,0,0,0,0,0)
id = OCOD("EXPP")
exitPupilPos = OPEV(id, 0,0,0,0,0,0)

surfCount = NSUR()

PRINT "# hjson format" 
PRINT "# postfixes: T = tangential, S = saggittal, im = image, obj = object"

str$ = "name: " + fName$
PRINT str$
str$ = "units: " + $UNITS()
PRINT str$

FORMAT 2 INT
PRINT "wavelength_count: ", waveCount
PRINT "primary_wavelength: ", primaryWave
PRINT "# in micrometers"
FORMAT 4.3
PRINT "wavelengths: ["
FOR i, 1, waveCount, 1
    str$ ="  "+ $STR(WAVL(i)) + ","
    PRINT str$
NEXT
PRINT "]"
PRINT

FORMAT 3 INT
PRINT "field_type: ", fieldTypeID
PRINT "# field types: "
PRINT "# 0 - degrees object space"
PRINT "# 1 - object heigth in lens units"
PRINT "# 2 - paraxial image height in lens units"
PRINT "# 3 - real image heigth in lens units"
PRINT "field_count: ", fieldCount
PRINT "# full half-field angle or height"
FORMAT 6.3
str$ = "max_field: " + $STR(maxField)
PRINT str$
! PRINT "unvignetted_field: ", unvignettedField
SETVIG #calculate vignetting 
PRINT "fields: ["
FOR i, 1, fieldCount, 1
    PRINT "  {"
    FORMAT 2 INT
    PRINT "    no: ", i
    FORMAT 9.6
    PRINT "    x_field                : ", FLDX(i)
    PRINT "    y_field                : ", FLDY(i)
    PRINT "    vignetting_angle       : ", FVAN(i)
    PRINT "    vignetting_compession_x: ", FVCX(i)
    PRINT "    vignetting_compession_y: ", FVCY(i)
    PRINT "    vignetting_decenter_x  : ", FVDX(i)
    PRINT "    vignetting_decenter_y  : ", FVDY(i)
    PRINT "  },"
NEXT
PRINT "]"
PRINT
PRINT "aperture_data: {"
FORMAT 6.3
PRINT "  type : ", apertureType
PRINT "  value: ", apertureValue
PRINT "  D_im : ", exitPupilDiam
PRINT "  D_obj: ", entrPupilDiam
str$ = "  ENPP: " + $STR(entrPupilPos) + "  #relative to first surface"
PRINT str$
str$ = "  EXPP: " + $STR(exitPupilPos) + " #relative to image surface"
PRINT str$
PRINT "}"
PRINT
FORMAT 3 INT
PRINT "surface_count: ", surfCount

!print surfaces
PRINT "# index is 0 for air in Zemax"
PRINT "surfaces: ["
FOR i, 1, surfCount, 1
    FORMAT 3 INT
    PRINT "  { no: ", i
    noreturn = SPRO(i, 1)
    str$ = $BUFFER()
    PRINT "    type: ", str$
    FORMAT 12.6 
    id = OCOD("POWR")
    PRINT "    power    : ", OPEV(id, i, primaryWave, 0, 0, 0, 0)
    PRINT "    curvature: ", CURV(i)
    PRINT "    thickness: ", THIC(i)
    PRINT "    conic    : ", CONI(i)
    PRINT "    edge     : ", EDGE(i)
    str$ = "    glass    : " + $GLASS(i)
    PRINT str$
    str$ = "    catalog  : " + $GLASSCATALOG(i)
    PRINT str$
    PRINT "    index@d  : ", GIND(i)
    PRINT "    abbe     : ", GABB(i)
    PRINT "  },"
NEXT
PRINT "]"
! end print surfaces
PRINT
PRINT
PRINT "#Aberrations are calculated for axial and two fields:"
PRINT "#1. axial"
PRINT "#2. full half-field w, linear vignetting k = 0.5"
PRINT "#3. unvignetted field, k = 1"
PRINT "#Aberrations are calculated for every wavelength."
PRINT
PRINT "maximum: {"
id = OCOD("DIMX")
! DIST(field, wave, absolute)
! 0 for max field
PRINT "    DIMX_percent: ", OPEV(id, 0, primaryWave, 0, 0, 0,0)
PRINT "}"
PRINT

!axial - transverse & longitudinal
! paraxial and for varying Py up to 1
! Py = 0 ; 0,5 ; 0,7 ; 1
! aberrations for all wavelengths
! offense against the sine cond for main wave
PRINT "axial: ["
FOR i, 1, Px_count, 1
    FOR j, 1, Py_count, 1
        PRINT "  {"
        FORMAT 5.4
        PRINT "  Px: ", Px(i)
        PRINT "  Py: ", Py(j)
        trax$ = ""
        tray$ = ""
        lona$ = ""
        FOR k, 1, waveCount, 1
            FORMAT 6.4 EXP
            IF afocal_im_space
                IF k > 1 
                    lona$ = lona$ + ", "
                ENDIF
                id = OCOD("ANAY")
                ! ANAY(void, wave, Hx, Hy, Px, Py)
                lona$ = lona$ + $STR(OPEV(id, 0, k, 0, 0, Px(i), Py(j)))
            ELSE
                IF k > 1 
                    lona$ = lona$ + ", "
                ENDIF
                id = OCOD("LONA")
                ! LONA(surface, wave, zone)
                lona$ = lona$ + $STR(OPEV(id, 0, k, Py(i), 0, 0, 0)) 
            ENDIF
            IF k > 1 
                trax$ = trax$ + ", "
                tray$ = tray$ + ", "
            ENDIF
            id = OCOD("TRAX")
            ! TRAY(surface, wave, Hx, Hy, Px, Py)
            trax$ = trax$ + $STR(OPEV(id, 0, k, 0, 0, Px(i), Py(j))) 
            id = OCOD("TRAY")
            ! TRAY(surface, wave, Hx, Hy, Px, Py)
            tray$ = tray$ + $STR(OPEV(id, 0, k, 0, 0, Px(i), Py(j))) 
        NEXT
        str$ = "  TRAX: [" + trax$ + "]"
        PRINT str$
        str$ = "  TRAY: [" + tray$ + "]"
        PRINT str$
        IF afocal_im_space
            str$ = "  ANAY: [" + lona$ + "]"
            PRINT str$
        ELSE
            str$ = "  LONA: [" + lona$ + "]"
            PRINT str$
        ENDIF
        id = OCOD("OSCD")
        ! OSCD(surface, wave, zone)
        PRINT "  OSCD:", OPEV(id, 0, primaryWave, Py(j), 0, 0, 0)
        PRINT "  },"
    NEXT
NEXT
PRINT "]"
PRINT

!chief ray - transverse only
!image size for all waves
!rest of parameters for the main wave
PRINT "chief: ["
FOR i, 1, Hx_count, 1
    FOR j, 1, Hy_count, 1
        PRINT "    {"
        PRINT "      Hx: ", Hx(i)
        PRINT "      Hy: ", Hy(j)
        FORMAT 6.4 
        IF afocal_im_space
            id = OCOD("RANG")
            rang$ = "      RANG: ["
            FOR wave, 1, waveCount, 1
                IF wave > 1
                    rang$ = rang$ + ", "
                ENDIF
                ! RANG(surface, wave, Hx, Hy, Px, Py)
                opval = OPEV(id, surfCount, wave, 0, 1, 0, 0)
                rang$ = rang$ + $STR(opval) 
            NEXT
            rang$ = rang$ + "]"
            PRINT rang$
        ELSE
            id = OCOD("REAX")
            reax$ = "      REAX: ["
            FOR wave, 1, waveCount, 1
                IF wave > 1
                    reax$ = reax$ + ", "
                ENDIF
                ! REAX(surface, wave, Hx, Hy, Px, Py)
                opval = OPEV(id, surfCount, wave, 0, 1, 0, 0)
                reax$ = reax$ + $STR(opval) + ", "
            NEXT
            reax$ = reax$ + "]"
            PRINT reax$ 
            id = OCOD("REAY")
            reay$ = "      REAY: ["
            FOR wave, 1, waveCount, 1
                IF wave > 1
                    reay$ = reay$ + ", "
                ENDIF
                ! REAY(surface, wave, Hx, Hy, Px, Py)
                opval = OPEV(id, surfCount, wave, 0, 1, 0, 0)
                reay$ = reay$ + $STR(opval) + ", "
            NEXT
            reay$ = reay$ + "]"
            PRINT reay$ 
        ENDIF
        id = OCOD("DISG")
        ! DISG(field, wave, Hx, Hy, Px, Py)
        str$ = "    DISG: " + $STR(OPEV(id, maxField, primaryWave, Hx(i), Hy(j), 0, 0)) + " #in %"
        PRINT str$
        PRINT "    },"
    NEXT
NEXT
PRINT "]"
!CONVERTFILEFORMAT jsonFilePath$, 1 !convert to ASCII
! cleanup so that there are no side effects
