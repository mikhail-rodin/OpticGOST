Attribute VB_Name = "optics"
Option Base 0
Option Explicit
Public Function SpectralLine(wavelength_nm As Double) As String
' returns a Fraunhofer line letter (e.g. e', F, G etc)
' for a wavelength in nanometers
'NB: all line symbols are 2 character wide for ease of formatting
    Dim fl As String
    Dim elt As String
    elt = ""
    fl = ""
    Select Case wavelength_nm
        Case 404 To 450
            fl = "h "
            elt = "Hg"
        Case 435 To 436
            fl = "g "
            elt = "Hg"
        Case 479 To 481
            fl = "F'"
            elt = "Cd"
        Case 486 To 487
            fl = "F "
            elt = "H"
        Case 545 To 546
            fl = "e "
            elt = "Hg"
        Case 587 To 588
            fl = "d "
            elt = "He"
        Case 589 To 590
            fl = "D "
            elt = "Na"
        Case 643 To 644
            fl = "C'"
            elt = "Cd"
        Case 656 To 657
            fl = "C "
            elt = "H"
        Case 706 To 707
            fl = "r "
            elt = "He"
        Case 852 To 853
            fl = "s "
            elt = "Cs"
        Case 1013 To 1014
            fl = "t "
            elt = "Hg"
    End Select
    SpectralLine = fl
End Function
