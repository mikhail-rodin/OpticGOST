! JSONexport
!
! This macro exports lens data (r,n,d, glass) and minimal
! aberration analysis data in a widespread easy-to-parse JSON format.
! Data can be imported into Excel to create GOST-compliant data sheets
! using the OpticGOST add-in. 
!
! The JSON file is saved alongside the zmx file. 
! A config.txt file has to be generated with a JSONconfig macro before exporting to JSON. 
!
! Pupil coords for calculating aberrations are set in config.txt. Thus you have the flexibility
! to set as many points on the pupil as you will.
!
! Aberrations are calculated for every field value set in Zemax: 
! if either Hx=0 or Hy=0 for a field (which's usually the case), X and Y axes
! correspond to tangential&sagittal, and isolated T&S aberrations will be calculated for it.
!
! https://github.com/mikhail-rodin/OpticGOST

PRINT 
PRINT "+---------------------------------------------+"
PRINT "|            OpticGOST v1.3.3                 |"
PRINT "| https://github.com/mikhail-rodin/OpticGOST  |"
PRINT "+---------------------------------------------+"
PRINT "|          JSON lens data export              |"
PRINT "+---------------------------------------------+"
PRINT
PRINT "Note: OpticGOST uses Zemax global coordinates,"
PRINT "so choosing a reference surface might make a difference."
PRINT
fileName$ = $FILENAME()
fNameLen = SLEN(fileName$)
!remove extension
fName$ = $LEFTSTRING(fileName$, fNameLen - 4)
!now fName in filename without exptension

zmxPath$ = $PATHNAME()
! where zmx file is stored

jsonFilePath$ = zmxPath$ + "\" + fName$ + "_lensdata.json"

! pupil coords are set in config file which's persed here
! Hx/Hy coords are calculated from field datal
DECLARE Hx, double, 1, 50
DECLARE Hy, double, 1, 50
DECLARE Px, double, 1, 50
DECLARE Py, double, 1, 50
! defaults
Py(1) = 0 
Py(2) = 0.5
Py(3) = 0.707
Py(4) = 1
Px(1) = 0
Px_count = 1
Py_count = 4

! these are coord vectors used to compute raytrace data
! vec = (axial, chief, upper, lower)^T
DECLARE raytrace_Hx, double, 1, 50
DECLARE raytrace_Hy, double, 1, 50
DECLARE raytrace_Px, double, 1, 50
DECLARE raytrace_Py, double, 1, 50
! these are fixed and not to be edited
! 1 - axial_y
raytrace_Px(1) = 0
raytrace_Py(1) = 1
raytrace_Hx(1) = 0
raytrace_Hy(1) = 0
! 2 - chief_y
raytrace_Px(2) = 0
raytrace_Py(2) = 0
raytrace_Hx(2) = 0
raytrace_Hy(2) = 1
! 3 - upper_y
raytrace_Px(3) = 0
raytrace_Py(3) = 1
raytrace_Hx(3) = 0
raytrace_Hy(3) = 1
! 4 - lower_y
raytrace_Px(4) = 0
raytrace_Py(4) = -1
raytrace_Hx(4) = 0
raytrace_Hy(4) = 1

! array of indices of coord sets for which raytrace data is calculated
DECLARE raytraceSelection, integer, 1, 8
! default selection - not to be changed
raytraceSelection(1) = 1
raytraceSelection(2) = 2
raytraceSelection(3) = 3
raytraceSelection(4) = 4
selectedRayCount = 4
! coord sets with index 5 and further are set in config file

! TODO: read raytrace coords from config

! TODO: fallback to defaults in case config can't be read
configFilePath$ = zmxPath$ + "\" + fName$ + "_config.txt"
OPEN configFilePath$
lineCounter = 0
PRINT "Reading export settings from"
PRINT configFilePath$
PRINT
LABEL 10
READSTRING confLine$
IF EOFF() == 1
    FORMAT 4 INT
    msg$ = "Config parsing completed, " + $STR(lineCounter) + " lines read."
    PRINT msg$
    PRINT 
    GOTO 20
ENDIF
lineCounter = lineCounter + 1
FORMAT 6.4
first$ = $GETSTRING(confLine$,1)
IF first$ $== "#" THEN GOTO 10
IF first$ $== "Px_count:" 
    str$ = $GETSTRING(confLine$,2)
    Px_count = SVAL(str$)
    GOTO 10
ELSE
    IF first$ $== "Py_count:" 
        str$ = $GETSTRING(confLine$,2)
        Py_count = SVAL(str$)
        GOTO 10
    ELSE
        IF first$ $== "Px:" 
            FOR i, 1, Px_count, 1
                str$ = $GETSTRING(confLine$,i+1)
                Px(i) = SVAL(str$)
            NEXT
            GOTO 10
        ELSE
            IF first$ $== "Py:" 
                FOR i, 1, Py_count, 1
                    str$ = $GETSTRING(confLine$,i+1)
                    Py(i) = SVAL(str$)
                NEXT
                GOTO 10
            ENDIF
        ENDIF
    ENDIF
ENDIF
GOTO 10
LABEL 20
CLOSE 

FORMAT 6.3
PRINT "Px (sagittal) coords from config:"
Px$ = ""
FOR i, 1, Px_count, 1
    Px$ = Px$ + " " + $STR(Px(i))
NEXT
PRINT Px$
PRINT "Py (tangential) coords from config:"
Py$ = ""
FOR i, 1, Py_count, 1
    Py$ = Py$ + " " + $STR(Py(i))
NEXT
PRINT Py$
PRINT
PRINT "Saving a JSON file with lens data and aberration analysis data" 
msg$ = "   to " + jsonFilePath$
PRINT msg$

OUTPUT jsonFilePath$
GETSYSTEMDATA 1
waveCount = NWAV()
primaryWave = PWAV()

afocal_im_space = SYPR(18)
id = OCOD("EFFL")
! EFFL(void, wave)
IF (afocal_im_space == 0) THEN effl = OPEV(id, 0, primaryWave, 0, 0, 0, 0)

fieldCount = NFLD()
fieldTypeID = SYPR(100)
! 0 for angle
! 1 for obj height
! 2 for paraxial image height
! 3 for real image height
maxField = MAXF()

apertureType = SYPR(10)
! 0 for entrance pupil diameter
! 1 for image space F/# ! 2 for object space num aperture NA ! 3 for float by stop size
! 4 for paraxial working F/#
! 5 for object cone angle in degrees
apertureValue = SYPR(11)
! EXPD(void)
id = OCOD("EXPD")
exitPupilDiam = OPEV(id, 0, 0, 0, 0, 0, 0)
IF (apertureType == 0) 
    entrPupilDiam = apertureValue
ELSE
    IF (apertureType == 1)
        entrPupilDiam = effl/apertureValue
    ELSE
        ! EPDI(void)
        id = OCOD("EPDI")         
        entrPupilDiam = OPEV(id, 0,0,0,0,0,0)
    ENDIF
ENDIF

id = OCOD("ENPP")
entrPupilPos = OPEV(id, 0,0,0,0,0,0)
id = OCOD("EXPP")
exitPupilPos = OPEV(id, 0,0,0,0,0,0)

surfCount = NSUR()
! surface numbers of first and last lens surfaces
firstSurf = 1
lastSurf = surfCount -1
imageSurf = surfCount

PRINT "# hjson format" 

str$ = "name: " + fName$
PRINT str$
str$ = "units: " + $UNITS()
PRINT str$
FORMAT 1 INT
PRINT "afocal: ", afocal_im_space
PRINT "telecentric_obj_space: ", VEC1(25)
FORMAT 5.2
PRINT "paraxial_magnification: ", VEC1(16)
PRINT "angular_magnification: ", VEC1(17)
PRINT "EFFL: ", VEC1(7)
PRINT "total_track: ", VEC1(18)
FORMAT 2 INT
PRINT "wavelength_count: ", waveCount
PRINT "primary_wavelength: ", primaryWave
FORMAT 4.3
wavelist$ = "wavelengths: ["
FOR i, 1, waveCount, 1
    IF i > 1
        wavelist$ = wavelist$ + ", "
    ENDIF
    wavelist$ = wavelist$ + $STR(WAVL(i)) 
NEXT
wavelist$ = wavelist$ + "]"
PRINT wavelist$
PRINT
PRINT "aperture_data: {"
FORMAT 1 INT
PRINT "  type : ", apertureType
FORMAT 6.3
PRINT "  value: ", apertureValue
PRINT "  WFNO : ", VEC1(10)
PRINT "  D_im : ", exitPupilDiam
PRINT "  D_obj: ", entrPupilDiam
str$ = "  ENPP : " + $STR(entrPupilPos) + "  #relative to first surface"
PRINT str$
str$ = "  EXPP : " + $STR(exitPupilPos) + " #relative to image surface"
PRINT str$
PRINT "}"
PRINT
FORMAT 3 INT
PRINT "surface_count: ", surfCount
PRINT "stop_surface_no: ", VEC1(23)
PRINT "# index is 0 for air in Zemax"
PRINT "# for each surface a chief ray trace data is computed"

PRINT "surfaces: ["
FOR i, 1, surfCount, 1
    FORMAT 3 INT
    PRINT "  { no       : ", i
    noreturn = SPRO(i, 1)
    str$ = $BUFFER()
    PRINT "    type     : ", str$
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
    FOR ray, 1, selectedRayCount, 1
        IF ray == 1 
            PRINT "    axial_y  : {"
        ELSE
            IF ray == 2
                PRINT "    chief_y  : {"
            ELSE
                IF ray == 3
                    PRINT "    upper_y  : {"
                ELSE
                    IF ray == 4
                        PRINT "    lower_y  : {"
                    ELSE
                        PRINT "    raytrace : {"
                        coords$ = "                coords: [" + $STR(raytrace_Hx(ray)) + ", " + $STR(raytrace_Hy(ray)) + ", " + $STR(raytrace_Px(ray)) + ", " + $STR(raytrace_Py(ray)) + "]"
                        PRINT coords$
                    ENDIF
                ENDIF
            ENDIF
        ENDIF
        raga$ = ""
        ragb$ = ""
        reax$ = ""
        reay$ = ""
        ssag$ = ""
        FOR k, 1, waveCount, 1
            IF k > 1
                raga$ = raga$ + ", "
                ragb$ = ragb$ + ", "
                reax$ = reax$ + ", "
                reay$ = reay$ + ", "
                ssag$ = ssag$ + ", "
            ENDIF
            id = OCOD("RAGA")
            raga$ = raga$ + $STR(OPEV(id,i,k,raytrace_Hx(ray),raytrace_Hy(ray),raytrace_Px(ray),raytrace_Py(ray)))
            id = OCOD("RAGB")
            ragb$ = ragb$ + $STR(OPEV(id,i,k,raytrace_Hx(ray),raytrace_Hy(ray),raytrace_Px(ray),raytrace_Py(ray)))
            id = OCOD("REAX")
            reax$ = reax$ + $STR(OPEV(id,i,k,raytrace_Hx(ray),raytrace_Hy(ray),raytrace_Px(ray),raytrace_Py(ray)))
            id = OCOD("REAY")
            reay$ = reay$ + $STR(OPEV(id,i,k,raytrace_Hx(ray),raytrace_Hy(ray),raytrace_Px(ray),raytrace_Py(ray)))
            id = OCOD("SSAG")
            ssag$ = ssag$ + $STR(OPEV(id,i,k,raytrace_Hx(ray),raytrace_Hy(ray),raytrace_Px(ray),raytrace_Py(ray)))
        NEXT
        raga$ = "            RAGA: [" + raga$ + "]"
        ragb$ = "            RAGB: [" + ragb$ + "]"
        reax$ = "            REAX: [" + reax$ + "]"
        reay$ = "            REAY: [" + reay$ + "]"
        ssag$ = "            SSAG: [" + ssag$ + "]"
        PRINT raga$
        PRINT ragb$
        PRINT reax$
        PRINT reay$
        PRINT ssag$
        PRINT "            }"
    NEXT
    PRINT "  },"
NEXT
PRINT "]"
PRINT
PRINT "maximum: {"
id = OCOD("DIMX")
! DIST(field, wave, absolute)
! 0 for max field
PRINT "    DIMX_percent: ", OPEV(id, 0, primaryWave, 0, 0, 0,0)
PRINT "}"
PRINT
PRINT "axial_x: ["
FOR coord, 1, Py_count, 1
    IF Py(coord) >= 0 
        PRINT "  {"
        FORMAT 5.4
        PRINT "  Px: ", Py(coord) 
        ! not a typo: Py coords are taken as tangential
        trax$ =      "  TRAX     : ["
        lona$ =      "  LONA     : ["
        anax$ =      "  ANAX     : ["
        entr_rang$ = "  entr_RANG: ["
        exit_rang$ = "  exit_RANG: ["
        h_1$ =       "  h_1      : ["
        h_q$ =       "  h_q      : ["
        FOR wave, 1, waveCount, 1
            FORMAT 6.3 EXP
            IF wave > 1 
                anax$ = anax$ + ", "
                trax$ = trax$ + ", "
                lona$ = lona$ + ", "
                entr_rang$ = entr_rang$ + ", "
                exit_rang$ = exit_rang$ + ", "
                h_1$ = h_1$ + ","
                h_q$ = h_q$ + ","
            ENDIF
            id = OCOD("ANAX")
            ! ANAY(void, wave, Hx, Hy, Px, Py)
            anax$ = anax$ + $STR(OPEV(id, 0, wave, 0, 0, Py(coord), 0))
            id = OCOD("TRAX")
            ! TRAX(surface, wave, Hx, Hy, Px, Py)
            trax$ = trax$ + $STR(OPEV(id, 0, wave, 0, 0, Py(coord), 0)) 
            id = OCOD("LONA")
            ! LONA(wave, void, zone)
            lona$ = lona$ + $STR(OPEV(id, wave, 0, Py(coord), 0, 0, 0)) 
            ! RANG(surface, wave, Hx, Hy, Px, Py)
            id = OCOD("RANG")
            entr_rang$ = entr_rang$ + $STR(OPEV(id,firstSurf,wave,0,0,Py(coord),0))
            exit_rang$ = exit_rang$ + $STR(OPEV(id,lastSurf,wave,0,0,Py(coord),0))
            id = OCOD("REAX")
            h_1$ = h_1$ + $STR(OPEV(id,firstSurf,wave,0,0,Py(coord),0))
            h_q$ = h_q$ + $STR(OPEV(id,lastSurf,wave,0,0,Py(coord),0))
        NEXT
        lona$ =  lona$ + "]"
        entr_rang$ =  entr_rang$ + "]"
        exit_rang$ =  + exit_rang$ + "]"
        anax$ =  + anax$ + "]"
        trax$ =  + trax$ + "]"
        h_1$ = h_1$ + "]"
        h_q$ = h_q$ + "]"
        PRINT lona$
        PRINT entr_rang$
        PRINT exit_rang$
        IF afocal_im_space
            PRINT anax$
        ELSE
            PRINT trax$
        ENDIF
        PRINT h_1$
        PRINT h_q$
        id = OCOD("OSCD")
        ! OSCD(surface, wave, zone)
        PRINT "  OSCD: ", OPEV(id, 0, primaryWave, Py(i), 0, 0, 0)
        PRINT "  },"
    ENDIF
NEXT
PRINT "]"
PRINT "axial_y: ["
FOR coord, 1, Py_count, 1
    IF Py(coord) >= 0 
        PRINT "  {"
        FORMAT 5.4
        PRINT "  Py: ", Py(coord) 
        ! not a typo: Py coords are taken as tangential
        tray$ =      "  TRAY     : ["
        lona$ =      "  LONA     : ["
        anay$ =      "  ANAY     : ["
        entr_rang$ = "  entr_RANG: ["
        exit_rang$ = "  exit_RANG: ["
        h_1$ =       "  h_1      : ["
        h_q$ =       "  h_q      : ["
        FOR wave, 1, waveCount, 1
            FORMAT 6.3 EXP
            IF wave > 1 
                anay$ = anay$ + ", "
                tray$ = tray$ + ", "
                lona$ = lona$ + ", "
                entr_rang$ = entr_rang$ + ", "
                exit_rang$ = exit_rang$ + ", "
                h_1$ = h_1$ + ","
                h_q$ = h_q$ + ","
            ENDIF
            id = OCOD("ANAY")
            ! ANAY(void, wave, Hx, Hy, Px, Py)
            anay$ = anay$ + $STR(OPEV(id, 0, wave, 0, 0, 0, Py(coord)))
            id = OCOD("TRAY")
            ! TRAY(surface, wave, Hx, Hy, Px, Py)
            tray$ = tray$ + $STR(OPEV(id, 0, wave, 0, 0, 0, Py(coord))) 
            id = OCOD("LONA")
            ! LONA(wave, void, zone)
            lona$ = lona$ + $STR(OPEV(id, wave, 0,Py(coord), 0, 0, 0)) 
            ! RANG(surface, wave, Hx, Hy, Px, Py)
            id = OCOD("RANG")
            entr_rang$ = entr_rang$ + $STR(OPEV(id,firstSurf,wave,0,0,0,Py(coord)))
            exit_rang$ = exit_rang$ + $STR(OPEV(id,lastSurf,wave,0,0,0,Py(coord)))
            id = OCOD("REAY")
            h_1$ = h_1$ + $STR(OPEV(id,firstSurf,wave,0,0,0,Py(coord)))
            h_q$ = h_q$ + $STR(OPEV(id,lastSurf,wave,0,0,0,Py(coord)))
        NEXT
        lona$ =  lona$ + "]"
        entr_rang$ =  entr_rang$ + "]"
        exit_rang$ =  + exit_rang$ + "]"
        anay$ =  + anay$ + "]"
        tray$ =  + tray$ + "]"
        h_1$ = h_1$ + "]"
        h_q$ = h_q$ + "]"
        PRINT lona$
        PRINT entr_rang$
        PRINT exit_rang$
        IF afocal_im_space
            PRINT anay$
        ELSE
            PRINT tray$
        ENDIF
        PRINT h_1$
        PRINT h_q$
        id = OCOD("OSCD")
        ! OSCD(surface, wave, zone)
        PRINT "  OSCD: ", OPEV(id, 0, primaryWave, Py(i), 0, 0, 0)
        PRINT "  },"
    ENDIF
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
PRINT "# full half-field angle or height"
FORMAT 6.3
str$ = "max_field: " + $STR(maxField)
PRINT str$
SETVIG 
PRINT "fields: ["
FOR field, 1, fieldCount, 1
    IF (FLDX(field)==0) & (FLDY(field)==0)
        GOTO 101
        ! next field w/o closing bracket
    ENDIF
    ! clockwise (positive) field angles are negative in Zemax
    ! so positive fields correspond to negative H coords
    Hx(field) = -FLDX(field)/maxField
    Hy(field) = -FLDY(field)/maxField
    PRINT "  {"
    FORMAT 2 INT
    PRINT "    no                     : ", field
    FORMAT 6.3
    PRINT "    Hx                     : ", Hx(field)
    PRINT "    Hy                     : ", Hy(field)
    PRINT "    x_field                : ", -FLDX(field)
    PRINT "    y_field                : ", -FLDY(field)
    PRINT "    vignetting_angle       : ", FVAN(field)
    PRINT "    vignetting_compression_x: ", FVCX(field)
    PRINT "    vignetting_compression_y: ", FVCY(field)
    PRINT "    vignetting_decenter_x  : ", FVDX(field)
    PRINT "    vignetting_decenter_y  : ", FVDY(field)
    ! chief ray aberrations
    FORMAT 6.3 EXP 
    rang$ =      "      RANG: ["
    entr_rang$ = " entr_RANG: ["
    h_1$ =       "       h_1: ["
    h_q$ =       "       h_q: ["
    FOR wave, 1, waveCount, 1
        IF wave > 1
            rang$ = rang$ + ", "
            entr_rang$ = entr_rang$ + ", "
            h_1$ = h_1$ + ", "
            h_q$ = h_q$ + ", "
        ENDIF
        ! to calculate real pupil positions, we need the following:
        ! 1. exit ray angle RANG(surface, wave, Hx, Hy, Px, Py)
        ! 2. chief ray height at last surface
        ! 3. entrance ray angle
        ! 4. chief ray height at 1st surface
        ! and for finite object space we also need image height. That's all.
        id = OCOD("RANG")
        rang$ = rang$ + $STR(OPEV(id, lastSurf, wave, Hx(field), Hy(field), 0, 0)) 
        entr_rang$ = entr_rang$ + $STR(OPEV(id, firstSurf, wave, Hx(field), Hy(field), 0, 0)) 
        IF FLDX(field) == 0
            id = OCOD("REAY")
        ELSE
            IF FLDY(field) == 0
                id = OCOD("REAX")
            ELSE 
                PRINT "  },"
                GOTO 101
                ! 101 = next field
            ENDIF
        ENDIF
        h_1$ = h_1$ + $STR(OPEV(id, 1, wave, Hx(field), Hy(field), 0, 0)) 
        h_q$ = h_q$ + $STR(OPEV(id,lastSurf, wave, Hx(field), Hy(field), 0, 0)) 
    NEXT
    rang$ = rang$ + "]"
    entr_rang$ = entr_rang$ + "]"
    h_1$ = h_1$ + "]"
    h_q$ = h_q$ + "]"
    PRINT "    chief: {"
    PRINT rang$
    PRINT entr_rang$
    PRINT h_1$
    PRINT h_q$
    LABEL 201
    IF afocal_im_space == 0 
        reax$ = "      REAX: ["
        reay$ = "      REAY: ["
        FOR wave, 1, waveCount, 1
            IF wave > 1
                reax$ = reax$ + ", "
                reay$ = reay$ + ", "
            ENDIF
            ! REAX(surface, wave, Hx, Hy, Px, Py)
            id = OCOD("REAX")
            reax$ = reax$ + $STR(OPEV(id, surfCount, wave, Hx(field), Hy(field), 0, 0)) 
            id = OCOD("REAY")
            reay$ = reay$ + $STR(OPEV(id, surfCount, wave, Hx(field), Hy(field), 0, 0)) 
        NEXT
        reax$ = reax$ + "]"
        reay$ = reay$ + "]"
        PRINT reax$ 
        PRINT reay$ 
    ENDIF
    id = OCOD("FCGS")
    str$ = "      FCGS: " + $STR(OPEV(id,0,primaryWave,Hx(field),Hy(field),0,0))
    PRINT str$
    id = OCOD("FCGT")
    str$ = "      FCGT: " + $STR(OPEV(id,0,primaryWave,Hx(field),Hy(field),0,0))
    PRINT str$
    id = OCOD("REAR")
    str$ = "      REAR: " + $STR(OPEV(id,0,primaryWave,Hx(field),Hy(field),0,0))
    PRINT str$
    id = OCOD("DISG")
    ! DISG(field, wave, Hx, Hy, Px, Py)
    str$ = "      DISG: " + $STR(OPEV(id, maxField, primaryWave, Hx(field), Hy(field), 0, 0)) 
    PRINT str$
    PRINT "    },"
    PRINT "    tangential: ["
    ! if Hy = 0, assume that tangential line is Py=0 (useful for anamorphic lenses etc)
    ! if Hx = 0 (usual case), find aberrations for varying Py
    FORMAT 6.3
    FOR coord, 1, Py_count, 1
        FORMAT 6.3
        IF FLDY(field) == 0
            PRINT   "        { Px: ", Py(coord)
        ELSE
            PRINT   "        { Py: ", Py(coord)
        ENDIF
        FORMAT 6.3 EXP
        trax$ = "        TRAX: ["
        tray$ = "        TRAY: ["
        anax$ = "        ANAX: ["
        anay$ = "        ANAY: ["
        FOR wave, 1, waveCount, 1
            IF wave > 1 
                trax$ = trax$ + ", "
                anax$ = anax$ + ", "
                tray$ = tray$ + ", "
                anay$ = anay$ + ", "
            ENDIF
            ! TRAX(surface, wave, Hx, Hy, Px, Py)
            id = OCOD("TRAX")
            opval = OPEV(id,surfCount,wave,Hx(field),0,Py(coord),0)
            trax$ = trax$ + $STR(opval)
            id = OCOD("ANAX")
            opval = OPEV(id,0,wave,Hx(field),0,Py(coord),0)
            anax$ = anax$ + $STR(opval)
            ! TRAY(surface, wave, Hx, Hy, Px, Py)
            id = OCOD("TRAY")
            opval = OPEV(id,surfCount,wave,0, Hy(field), 0, Py(coord))
            tray$ = tray$ + $STR(opval)
            ! ANAY(void, wave, Hx, Hy, Px, Py)
            id = OCOD("ANAY")
            opval = OPEV(id,0,wave,0, Hy(field), 0, Py(coord))
            anay$ = anay$ + $STR(opval)
        NEXT
        tray$ = tray$ + "]"
        anay$ = anay$ + "]"
        trax$ = trax$ + "]"
        anax$ = anax$ + "]"
        IF afocal_im_space
            IF FLDY(field) == 0
                PRINT anax$
            ELSE
                PRINT anay$
            ENDIF
        ELSE
            IF FLDY(field) == 0
                PRINT trax$
            ELSE
                PRINT tray$
            ENDIF
        ENDIF
        PRINT "        },"
    NEXT 
    PRINT "    ]"
    PRINT "    sagittal: ["
    FOR coord, 1, Px_count, 1
        IF Px(coord) > 0 
            FORMAT 6.3
            IF FLDY(field) == 0
                PRINT   "        { Py: ", Px(coord)
            ELSE
                PRINT   "        { Px: ", Px(coord)
            ENDIF
            FORMAT 6.3 EXP
            trax$ = "        TRAX: ["
            tray$ = "        TRAY: ["
            anax$ = "        ANAX: ["
            anay$ = "        ANAY: ["
            FOR wave, 1, waveCount, 1
                IF wave > 1 
                    trax$ = trax$ + ", "
                    tray$ = tray$ + ", "
                    anax$ = anax$ + ", "
                    anay$ = anay$ + ", "
                ENDIF
                IF FLDY(field) == 0
                    ! TRAX(surface, wave, Hx, Hy, Px, Py)
                    id = OCOD("TRAX")
                    opval = OPEV(id,surfCount,wave,Hx(field),0,0,Px(coord))
                    trax$ = trax$ + $STR(opval)
                    id = OCOD("TRAY")
                    opval = OPEV(id,surfCount,wave,Hx(field),0,0,Px(coord))
                    tray$ = tray$ + $STR(opval)
                    ! ANAX(void, wave, Hx, Hy, Py, Py)
                    id = OCOD("ANAX")
                    opval = OPEV(id,surfCount,wave,Hx(field),0,0,Px(coord))
                    anax$ = anax$ + $STR(opval)
                    id = OCOD("ANAY")
                    opval = OPEV(id,surfCount,wave,Hx(field),0,0,Px(coord))
                    anay$ = anay$ + $STR(opval)
                ELSE
                    ! TRAX(surface, wave, Hx, Hy, Px, Py)
                    id = OCOD("TRAX")
                    opval = OPEV(id,surfCount,wave,0,Hy(field),Px(coord),0)
                    trax$ = trax$ + $STR(opval)
                    id = OCOD("TRAY")
                    opval = OPEV(id,surfCount,wave,0,Hy(field),Px(coord),0)
                    tray$ = tray$ + $STR(opval)
                    ! ANAX(void, wave, Hx, Hy, Px, Py)
                    id = OCOD("ANAX")
                    opval = OPEV(id,surfCount,wave,0,Hy(field),Px(coord),0)
                    anax$ = anax$ + $STR(opval)
                    id = OCOD("ANAY")
                    opval = OPEV(id,surfCount,wave,0,Hy(field),Px(coord),0)
                    anay$ = anay$ + $STR(opval)
                ENDIF
            NEXT
            trax$ = trax$ + "]"
            tray$ = tray$ + "]"
            anax$ = anax$ + "]"
            anay$ = anay$ + "]"
            IF afocal_im_space
                PRINT anax$
                PRINT anay$
            ELSE
                PRINT trax$
                PRINT tray$
            ENDIF
            PRINT    "        },"
        ENDIF
    NEXT
    PRINT "    ]"
    PRINT "  },"
    LABEL 101
NEXT
PRINT "]"
