Contents of HJSON lensdata file.
Formatted as (key), (type), (units)

Top level parameters:
1. General system parameters
    1. name, string
    1. units, string
    1. wavelength_count, int
    1. primary_wavelength, int, micrometers
    1. field_type, int
    1. field_count, int
    1. max_field, float, field units
    1. unvignetted_field, float, field units
    1. surface_count, int
1. wavelengths, float[nWaves], in micrometers
1. fields, field_t[nFields]
1. aperture_data, aperture_t
1. surfaces, surface_t[nSurfaces]
1. maximum, maxAber_t #maximum aberrations
1. axial, axialAber_t[nCoords] #axial ray aberrations
1. chief, chiefAberSize_t[nWaves]

Types:
1. type field_t
    1. no, int
    1. x_field, float
    1. y_field, float
    1. vignetting_angle, float
    1. vignetting_compession_x, float
    1. vignetting_compession_y, float
    1. vignetting_decenter_x, float
    1. vignetting_decenter_y, float
1. type aperture_t
    1. type, int
    1. value, float
    1. D_im, float
    1. D_obj, float
    1. ENPP, float   #relative to first surface
    1. EXPP, float  #relative to image surface
1. type surface_t
    1. no, int
    1. type, int
    1. power, float
    1. curvature, float
    1. thickness, float
    1. conic, float
    1. edge, float
    1. glass, string
    1. catalog, string
    1. index_d, float
    1. abbe, float
1. maxAber_t
    1. DIMX_percent, float
1. axialAber_t
    1. Px, float
    1. Py, float
    1. TRAX, float[nWaves]
    1. TRAY, float[nWaves]
    1. ANAY/LONA, float[nWaves]
    1. OSCD, float
1. chiefAberSize_t
    1. Hx, float
    1. Hy, float
    1. REAX, float[nWaves]
    1. REAY, float[nWaves]
    1. DISG, float[nWaves]

     