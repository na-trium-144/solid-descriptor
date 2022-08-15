Attribute VB_Name = "SelType"
Sub ExportEntityName(SelType As Integer, swEntRef As Object, ByRef Name As String, ByRef ID() As String)
Name = ""
ID(0) = ""
ID(1) = ""
Select Case SelType
    Case swSelEDGES, swSelFACES, swSelVERTICES
        'No Name
    Case swSelDATUMPLANES, swSelDATUMAXES, swSelDATUMPOINTS
        Name = swEntRef.Name
    Case swSelSKETCHSEGS, swSelSKETCHPOINTS, swSelEXTSKETCHSEGS, swSelEXTSKETCHPOINTS
        Name = swEntRef.GetSketch().Name
        ID(0) = swEntRef.GetID()(0)
        ID(1) = swEntRef.GetID()(1)
    Case swSelCOMPONENTS
        Name = swEntRef.Name2
        'Do Until swEntRef.GetParent() Is Nothing
        '    Set swEntRef = swEntRef.GetParent()
        '    Name = Name & "@" & swEntRef.Name2
        'Loop
    Case Else
        MsgBox "Unsupported type " & SelType
End Select
End Sub
Function GetSelTypeString(SelType As Integer) As String
Select Case SelType
'Case    swSelNOTHING   :   GetSelTypeString =
Case swSelEDGES:       GetSelTypeString = "EDGE"
Case swSelFACES:       GetSelTypeString = "FACE"
Case swSelVERTICES:         GetSelTypeString = "VERTEX"
Case swSelDATUMPLANES:          GetSelTypeString = "PLANE"
Case swSelDATUMAXES:        GetSelTypeString = "AXIS"
Case swSelDATUMPOINTS:          GetSelTypeString = "DATUMPOINT"
Case swSelOLEITEMS:         GetSelTypeString = "OLEITEM"
Case swSelATTRIBUTES:       GetSelTypeString = "ATTRIBUTE"
Case swSelSKETCHES:         GetSelTypeString = "SKETCH"
Case swSelSKETCHSEGS:       GetSelTypeString = "SKETCHSEGMENT"
Case swSelSKETCHPOINTS:        GetSelTypeString = "SKETCHPOINT"
Case swSelDRAWINGVIEWS:         GetSelTypeString = "DRAWINGVIEW"
Case swSelGTOLS:        GetSelTypeString = "GTOL"
Case swSelDIMENSIONS:       GetSelTypeString = "DIMENSION"
Case swSelNOTES:        GetSelTypeString = "NOTE"
Case swSelSECTIONLINES:         GetSelTypeString = "SECTIONLINE"
Case swSelDETAILCIRCLES:        GetSelTypeString = "DETAILCIRCLE"
Case swSelSECTIONTEXT:          GetSelTypeString = "SECTIONTEXT"
Case swSelSHEETS:       GetSelTypeString = "SHEET"
Case swSelCOMPONENTS:          GetSelTypeString = "COMPONENT"
Case swSelMATES:        GetSelTypeString = "MATE"
Case swSelBODYFEATURES:         GetSelTypeString = "BODYFEATURE"
Case swSelREFCURVES:        GetSelTypeString = "REFCURVE"
Case swSelEXTSKETCHSEGS:        GetSelTypeString = "EXTSKETCHSEGMENT"
Case swSelEXTSKETCHPOINTS:          GetSelTypeString = "EXTSKETCHPOINT"
Case swSelHELIX:       GetSelTypeString = "HELIX"
Case swSelREFERENCECURVES:          GetSelTypeString = "REFERENCECURVES"
Case swSelREFSURFACES:         GetSelTypeString = "REFSURFACE"
Case swSelCENTERMARKS:          GetSelTypeString = "CENTERMARKS"
Case swSelINCONTEXTFEAT:        GetSelTypeString = "INCONTEXTFEAT"
Case swSelMATEGROUP:       GetSelTypeString = "MATEGROUP"
Case swSelBREAKLINES:          GetSelTypeString = "BREAKLINE"
Case swSelINCONTEXTFEATS:       GetSelTypeString = "INCONTEXTFEATS"
Case swSelMATEGROUPS:       GetSelTypeString = "MATEGROUPS"
Case swSelSKETCHTEXT:       GetSelTypeString = "SKETCHTEXT"
Case swSelSFSYMBOLS:       GetSelTypeString = "SFSYMBOL"
Case swSelDATUMTAGS:       GetSelTypeString = "DATUMTAG"
Case swSelCOMPPATTERN:         GetSelTypeString = "COMPPATTERN"
Case swSelWELDS:       GetSelTypeString = "WELD"
Case swSelCTHREADS:        GetSelTypeString = "CTHREAD"
Case swSelDTMTARGS:        GetSelTypeString = "DTMTARG"
Case swSelPOINTREFS:        GetSelTypeString = "POINTREF"
Case swSelDCABINETS:       GetSelTypeString = "DCABINET"
Case swSelEXPLVIEWS:       GetSelTypeString = "EXPLODEDVIEWS"
Case swSelEXPLSTEPS:       GetSelTypeString = "EXPLODESTEPS"
Case swSelEXPLLINES:       GetSelTypeString = "EXPLODELINES"
Case swSelSILHOUETTES:         GetSelTypeString = "SILHOUETTE"
Case swSelCONFIGURATIONS:          GetSelTypeString = "CONFIGURATIONS"
'Case    swSelOBJHANDLES    :   GetSelTypeString =
Case swSelARROWS:       GetSelTypeString = "VIEWARROW"
Case swSelZONES:       GetSelTypeString = "ZONES"
Case swSelREFEDGES:        GetSelTypeString = "REFERENCE-EDGE"
'Case    swSelREFFACES  :   GetSelTypeString =
'Case    swSelREFSILHOUETTE :   GetSelTypeString =
Case swSelBOMS:        GetSelTypeString = "BOM"
Case swSelEQNFOLDER:       GetSelTypeString = "EQNFOLDER"
Case swSelSKETCHHATCH:          GetSelTypeString = "SKETCHHATCH"
Case swSelIMPORTFOLDER:        GetSelTypeString = "IMPORTFOLDER"
Case swSelVIEWERHYPERLINK:          GetSelTypeString = "HYPERLINK"
'Case    swSelMIDPOINTS :   GetSelTypeString =
Case swSelCUSTOMSYMBOLS - Obsolete:       GetSelTypeString = "CUSTOMSYMBOL"
Case swSelCOORDSYS:        GetSelTypeString = "COORDSYS"
Case swSelDATUMLINES:          GetSelTypeString = "REFLINE"
'Case    swSelROUTECURVES   :   GetSelTypeString =
Case swSelBOMTEMPS:         GetSelTypeString = "BOMTEMP"
Case swSelROUTEPOINTS:         GetSelTypeString = "ROUTEPOINT"
Case swSelCONNECTIONPOINTS:        GetSelTypeString = "CONNECTIONPOINT"
'Case    swSelROUTESWEEPS   :   GetSelTypeString =
Case swSelPOSGROUP:        GetSelTypeString = "POSGROUP"
Case swSelBROWSERITEM:         GetSelTypeString = "BROWSERITEM"
Case swSelFABRICATEDROUTE:          GetSelTypeString = "ROUTEFABRICATED"
Case swSelSKETCHPOINTFEAT:        GetSelTypeString = "SKETCHPOINTFEAT"
'Case    swSelCOMPSDONTOVERRIDE  :   GetSelTypeString =
Case swSelLIGHTS:       GetSelTypeString = "LIGHTS"
'Case    swSelWIREBODIES :   GetSelTypeString =
Case swSelSURFACEBODIES:        GetSelTypeString = "SURFACEBODY"
Case swSelSOLIDBODIES:          GetSelTypeString = "SOLIDBODY"
Case swSelFRAMEPOINT:       GetSelTypeString = "FRAMEPOINT"
'Case    swSelSURFBODIESFIRST    :   GetSelTypeString =
Case swSelMANIPULATORS:         GetSelTypeString = "MANIPULATOR"
Case swSelPICTUREBODIES:        GetSelTypeString = "PICTURE BODY"
'Case    swSelSOLIDBODIESFIRST   :   GetSelTypeString =
Case swSelLEADERS:          GetSelTypeString = "LEADER"
Case swSelSKETCHBITMAP:         GetSelTypeString = "SKETCHBITMAP"
Case swSelDOWELSYMS:        GetSelTypeString = "DOWLELSYM"
Case swSelEXTSKETCHTEXT:        GetSelTypeString = "EXTSKETCHTEXT"
Case swSelBLOCKINST - Obsolete:        GetSelTypeString = "BLOCKINST"
Case swSelFTRFOLDER:        GetSelTypeString = "FTRFOLDER"
Case swSelSKETCHREGION:         GetSelTypeString = "SKETCHREGION"
Case swSelSKETCHCONTOUR:        GetSelTypeString = "SKETCHCONTOUR"
Case swSelBOMFEATURES:          GetSelTypeString = "BOMFEATURE"
Case swSelANNOTATIONTABLES:         GetSelTypeString = "ANNOTATIONTABLES"
Case swSelBLOCKDEF:         GetSelTypeString = "BLOCKDEF"
Case swSelCENTERMARKSYMS:       GetSelTypeString = "CENTERMARKSYMS"
Case swSelSIMULATION:       GetSelTypeString = "SIMULATION"
Case swSelSIMELEMENT:       GetSelTypeString = "SIMULATION_ELEMENT"
Case swSelCENTERLINES:          GetSelTypeString = "CENTERLINE"
Case swSelHOLETABLEFEATS:       GetSelTypeString = "HOLETABLE"
Case swSelHOLETABLEAXES:        GetSelTypeString = "HOLETABLEAXIS"
Case swSelWELDMENT:         GetSelTypeString = "WELDMENT"
Case swSelSUBWELDFOLDER:        GetSelTypeString = "SUBWELDMENT"
'Case    swSelEXCLUDEMANIPULATORS    :   GetSelTypeString =
Case swSelREVISIONTABLE:        GetSelTypeString = "REVISIONTABLE"
Case swSelSUBSKETCHINST:        GetSelTypeString = "SUBSKETCHINST"
Case swSelWELDMENTTABLEFEATS:       GetSelTypeString = "WELDMENTTABLE"
Case swSelBODYFOLDER:       GetSelTypeString = "BDYFOLDER"
Case swSelREVISIONTABLEFEAT:        GetSelTypeString = "REVISIONTABLEFEAT"
'Case    swSelSUBATOMFOLDER  :   GetSelTypeString =
Case swSelWELDBEADS3:       GetSelTypeString = "WELDBEADS"
Case swSelEMBEDLINKDOC:         GetSelTypeString = "EMBEDLINKDOC"
Case swSelJOURNAL:          GetSelTypeString = "JOURNAL"
Case swSelDOCSFOLDER:       GetSelTypeString = "DOCSFOLDER"
Case swSelCOMMENTSFOLDER:       GetSelTypeString = "COMMENTSFOLDER"
Case swSelCOMMENT:          GetSelTypeString = "COMMENT"
Case swSelCAMERAS:          GetSelTypeString = "CAMERAS"
Case swSelMATESUPPLEMENT:       GetSelTypeString = "MATESUPPLEMENT"
Case swSelANNOTATIONVIEW:       GetSelTypeString = "ANNVIEW"
Case swSelGENERALTABLEFEAT:         GetSelTypeString = "GENERALTABLEFEAT"
Case swSelSUBSKETCHDEF:         GetSelTypeString = "SUBSKETCHDEF"
Case swSelDISPLAYSTATE:         GetSelTypeString = "VISUALSTATE"
Case swSelTITLEBLOCK:       GetSelTypeString = "TITLEBLOCK"
Case swSelEVERYTHING:          GetSelTypeString = "EVERYTHING"
Case swSelLOCATIONS:        GetSelTypeString = "LOCATIONS"
Case swSelUNSUPPORTED:          GetSelTypeString = "UNSUPPORTED"
Case swSelSWIFTANNOTATIONS:         GetSelTypeString = "SWIFTANN"
Case swSelSWIFTFEATURES:        GetSelTypeString = "SWIFTFEATURE"
Case swSelSWIFTSCHEMA:          GetSelTypeString = "SWIFTSCHEMA"
Case swSelTITLEBLOCKTABLEFEAT:          GetSelTypeString = "TITLEBLOCKTABLEFEAT"
Case swSelOBJGROUP:         GetSelTypeString = "OBJGROUP"
Case swSelCOSMETICWELDS:        GetSelTypeString = "COSMETICWELDS"
Case SwSelMAGNETICLINES:        GetSelTypeString = "MAGNETICLINES"
Case swSelSELECTIONSETFOLDER:       GetSelTypeString = "SELECTIONSETFOLDER"
Case swSelSELECTIONSETNODE:         GetSelTypeString = "SUBSELECTIONSETNODE"
Case swSelPUNCHTABLEFEATS:          GetSelTypeString = "PUNCHTABLE"
Case swSelHOLESERIES:       GetSelTypeString = "HOLESERIES"
Case Else
    MsgBox "Unsupported type " & SelType
End Select
End Function
