Attribute VB_Name = "optics"
Option Base 0
Option Explicit
Function SpectralLine(wavelength_nm As Double) As String
' returns a Fraunhofer line letter (e.g. e', F, G etc)
' for a wavelength in nanometers
    Select Case wavelength_nm
    Case 888.8 To 908.8
        SpectralLine = "y"
    Case 812.7 To 832.7
        SpectralLine = "Z"
    Case 749.4 To 769.4
        SpectralLine = "A"
    Case 676.7 To 696.7
        SpectralLine = "B"
    Case 646.3 To 666.3
        SpectralLine = "C"
    Case 617.7 To 637.7
        SpectralLine = "D"
    Case 579.6 To 599.6
        SpectralLine = "D1"
    Case 579# To 599#
        SpectralLine = "D2"
    Case 577.6 To 597.6
        SpectralLine = "D3"
    Case 536.1 To 556.1
        SpectralLine = "e"
    Case 517# To 537#
        SpectralLine = "E2"
    Case 508.4 To 528.4
        SpectralLine = "b1"
    Case 507.3 To 527.3
        SpectralLine = "b2"
    Case 506.9 To 526.9
        SpectralLine = "b3"
    Case 506.7 To 526.7
        SpectralLine = "b4"
    Case 485.8 To 505.8
        SpectralLine = "c"
    Case 476.1 To 496.1
        SpectralLine = "F"
    Case 456.8 To 476.8
        SpectralLine = "d"
    Case 428.4 To 448.4
        SpectralLine = "e'"
    Case 424# To 444#
        SpectralLine = "G'"
    Case 420.8 To 440.8
        SpectralLine = "G"
    Case 420.8 To 440.8
        SpectralLine = "g"
    Case 400.2 To 420.2
        SpectralLine = "h"
    Case 386.8 To 406.8
        SpectralLine = "H"
    Case 383.4 To 403.4
        SpectralLine = "K"
    Case 372# To 392#
        SpectralLine = "L"
    Case 348.1 To 368.1
        SpectralLine = "N"
    Case 326.1 To 346.1
        SpectralLine = "P"
    Case 292.1 To 312.1
        SpectralLine = "T"
    Case 289.4 To 309.4
        SpectralLine = "t"
    End Select
End Function
