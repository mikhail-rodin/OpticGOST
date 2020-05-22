! export JSON-formatted lens data

fileName$ = $FILENAME()
fNameLen = SLEN(fileName$)
!remove extension
fName$ = $LEFTSTRING(fileName$, fNameLen - 4)
!now fName in filename without exptension

zmxPath$ = $PATHNAME()
! where zmx file is stored

jsonFilePath$ = zmxPath$ + "\" + fName$ + "_lensdata.json"

! there are no enum or lists in ZPL
! so we use an array to store the pupil coordinates we'll be calculating aberrations at
pupilCoordsDim = 4
DECLARE pupilCoords, double, 1, pupilCoordsDim
pupilCoords(1) = 0 
pupilCoords(2) = 0.5
pupilCoords(3) = 0.7
pupilCoords(4) = 1

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
        !IF (apertureType == 2)
            ! NA = sin (arctg(D/2f))
            ! arcsin NA = arctg D/2f
            ! tg arcsin NA = D/2f
        id = OCOD("EPDI")
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
PRINT "{"

str$ = "name: " + fName$
PRINT str$

FORMAT 2 INT
PRINT "wavelength_count: ", waveCount
PRINT "primary_wave_no: ", primaryWave
PRINT "# in micrometers"
FORMAT 4.3
PRINT "wavelengths: ["
FOR i, 1, waveCount, 1
    PRINT "  ", WAVL(i)
NEXT
PRINT "]"
PRINT

! find unvignetted half-field
! beta version: promt for it
INPUT "Unvignetted field: ", unvignettedField

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
PRINT "unvignetted_field: ", unvignettedField
SETVIG #calculate vignetting 
PRINT "fields: ["
FOR i, 1, fieldCount, 1
    PRINT "  {"
    FORMAT 2 INT
    PRINT "    no: ", i
    FORMAT 9.6
    PRINT "    x_field: ", FLDX(i)
    PRINT "    y_field: ", FLDY(i)
    PRINT "    vignetting_angle: ", FVAN(i)
    PRINT "    vignetting_compession_x: ", FVCX(i)
    PRINT "    vignetting_compession_y: ", FVCY(i)
    PRINT "    vignetting_decenter_x: ", FVDX(i)
    PRINT "    vignetting_decenter_y: ", FVDY(i)
    PRINT "  }"
NEXT
PRINT "]"
PRINT
PRINT "aperture_data: {"
FORMAT 6.3
PRINT "  type: ", apertureType
PRINT "  value: ", apertureValue
PRINT "  D_im: ", exitPupilDiam
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
i = 0
FOR i, 1, surfCount, 1
    FORMAT 3 INT
    PRINT "  { no: ", i
    noreturn = SPRO(i, 1)
    str$ = $BUFFER()
    PRINT "    type: ", str$
    FORMAT 12.6 
    id = OCOD("POWR")
    PRINT "    power: ", OPEV(id, i, primaryWave, 0, 0, 0, 0)
    PRINT "    curvature: ", CURV(i)
    PRINT "    thickness: ", THIC(i)
    PRINT "    conic: ", CONI(i)
    PRINT "    edge: ", EDGE(i)
    str$ = "    glass: " + $GLASS(i)
    PRINT str$
    str$ = "    catalog: " + $GLASSCATALOG(i)
    PRINT str$
    PRINT "    index@d: ", GIND(i)
    PRINT "    abbe: ", GABB(i)
    PRINT "  }"
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
FOR i, 1, pupilCoordsDim, 1
    PRINT "  {"
    FORMAT 5.4
    PRINT "  Py: ", pupilCoords(i)
    PRINT "  aberrations: ["
    FOR j, 1, waveCount, 1
        FORMAT 2 INT
        PRINT "    { wave :", j
        FORMAT 6.4 EXP
        IF afocal_im_space
            id = OCOD("ANAY")
            ! ANAY(void, wave, Hx, Hy, Px, Py)
            PRINT "      ANAY: ", OPEV(id, 0, j, 0, 0, 0, pupilCoords(i))
        ELSE
            id = OCOD("TRAY")
            ! TRAY(surface, wave, Hx, Hy, Px, Py)
            PRINT "      TRAY: ", OPEV(id, 0, j, 0, 0, 0, pupilCoords(i))
        ENDIF
        id = OCOD("LONA")
        ! LONA(surface, wave, zone)
        PRINT "      LONA: ", OPEV(id, 0, j, pupilCoords(i), 0, 0, 0)
        PRINT "    }"
    NEXT
    PRINT "  ]"
    id = OCOD("OSCD")
    ! OSCD(surface, wave, zone)
    PRINT "  OSCD:", OPEV(id, 0, primaryWave, pupilCoords(i), 0, 0, 0)
    PRINT "  }"
NEXT
PRINT "]"
PRINT

!chief ray - transverse only
!image size for all waves
!rest of parameters for the main wave
PRINT "chief: {"

PRINT "  max_field: {"
PRINT "    m_im: "
PRINT "    image_size: ["
FOR wave, 1, waveCount, 1
    FORMAT 1 INT
    PRINT "    { wave: ", wave
    FORMAT 6.4 
    IF afocal_im_space
        id = OCOD("RANG")
        ! RANG(surface, wave, Hx, Hy, Px, Py)
        PRINT "      RANG: ", OPEV(id, surfCount, wave, 0, 1, 0, 0)
    ELSE
        id = OCOD("REAY")
        ! REAY(surface, wave, Hx, Hy, Px, Py)
        PRINT "      REAY: ", OPEV(id, surfCount, wave, 0, 1, 0, 0)
    ENDIF
    PRINT "    }"
NEXT
PRINT "    ]"
FORMAT 4.2
id = OCOD("DISG")
! DISG(field, wave, Hx, Hy, Px, Py)
str$ = "    DISG: " + $STR(OPEV(id, maxField, primaryWave, 0, 1, 0, 0)) + " #in %"
PRINT str$
PRINT "  }"
PRINT "  unvignetted_field: {"
IF (fieldTypeID == 0)
    ! if field is angles in obj space
    Hy = TANG(unvignettedField)/TANG(maxField)
ELSE
    Hy = unvignettedField/maxField
ENDIF
PRINT "    image_size: ["
FOR wave, 1, waveCount, 1
    FORMAT 1 INT
    PRINT "    { wave: ", wave
    FORMAT 6.4
    IF afocal_im_space
        id = OCOD("RANG")
        ! RANG(surface, wave, Hx, Hy, Px, Py)
        PRINT "      RANG: ", OPEV(id, surfCount, wave, 0, Hy, 0, 0)
    ELSE
        id = OCOD("REAY")
        ! REAY(surface, wave, Hx, Hy, Px, Py)
        PRINT "      REAY: ", OPEV(id, surfCount, wave, 0, Hy, 0, 0)
    ENDIF
    PRINT "    }"
NEXT
PRINT "    ]"
FORMAT 4.2
id = OCOD("DISG")
! DISG(field, wave, Hx, Hy, Px, Py)
str$ = "    DISG: " + $STR(OPEV(id, unvignettedField, primaryWave, 0, Hy, 0, 0)) + " #in %"
PRINT str$

PRINT "  }"

!tangential - transverse only

!saggittal - transverse only

PRINT "}"

! cleanup so that there are no side effects
