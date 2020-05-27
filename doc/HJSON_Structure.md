Contents of HJSON lensdata file.
Formatted as (key), (type), (units)

Top level parameters:
1. General system parameters
    1. name, string
    1. wavelength_count, int
    1. primary_wavelength, int, micrometers
    1. field_type, int
    1. field_count, int
    1. max_field, float, field units
    1. unvignetted_field, float, field units
    1. surface_count, int
1. Parameters used by parser
    1. Py_coord_count, int 
1. wavelengths, float[], in micrometers
1. fields, fieldObject[]
1. aperture_data, apertureObject
1. surfaces, surfaceObject[]
1. maximum, maxAberObject #maximum aberrations
1. axial, axialAberObject[] #axial ray aberrations
1. chief, {max_field, unvignetted_field}
    1. max_field, chiefAberObject
    1. unvignetted_field, chiefAberObject


Types:
1. type fieldObject
    1. no, int
    1. x_field, float
    1. y_field, float
    1. vignetting_angle, float
    1. vignetting_compession_x, float
    1. vignetting_compession_y, float
    1. vignetting_decenter_x, float
    1. vignetting_decenter_y, float
1. type apertureObject
    1. type, int
    1. value, float
    1. D_im, float
    1. D_obj, float
    1. ENPP, float   #relative to first surface
    1. EXPP, float  #relative to image surface
1. type surfaceObject
    1. no, int
    1. type, int
    1. power, float
    1. curvature, float
    1. thickness, float
    1. conic, float
    1. edge, float
    1. glass, string
    1. catalog, string
    1. index@d, float
    1. abbe, float
1. maxAberObject
    1. DIMX_percent, float
1. axialAberObject
    1. Py, float
    1. aberrations, aberObject[]
    1. OSCD, float
1. aberObject
    1. wave, int
    1. TRAY, float
    1. LONA, float
1. chiefAberObject
    1. image_size, imSizeObject
    1. DISG, float
1. imSizeObject
    1. wave, int
    1. REAY, float

     