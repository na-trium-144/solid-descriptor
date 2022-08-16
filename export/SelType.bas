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
    Case Else
        MsgBox "Unsupported type " & SelType
End Select
End Sub

Sub HideEntityComponent(mtEntNode As IXMLDOMElement, swAsmDoc As IAssemblyDoc)

Dim SelectState As Boolean
SelectState = False
Dim SelectName As String
Dim ComponentName As String
ComponentName = XMLImport.cpNameReplace(mtEntNode.getAttribute("component"))

Dim swSelMgr As ISelectionMgr
Set swSelMgr = swAsmDoc.SelectionManager
Dim swSelectData As ISelectData
Set swSelectData = swSelMgr.CreateSelectData()
swSelectData.Mark = 1

swSelMgr.SuspendSelectionList

Dim swRefCp As IComponent2
SelectName = ComponentName
SelectState = swAsmDoc.Extension.SelectByID2(SelectName, "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
Debug.Print SelectName
Debug.Print SelectState
If Not SelectState Then Exit Sub

Set swRefCp = swSelMgr.GetSelectedObject6(swSelMgr.GetSelectedObjectCount2(1), 1)

swSelMgr.ResumeSelectionList2 False


Dim swRefCp_p As IComponent2
Set swRefCp_p = swRefCp
Do Until swRefCp_p Is Nothing
    swRefCp_p.SetVisibility swComponentHidden, swThisConfiguration, ""
    Set swRefCp_p = swRefCp_p.GetParent()
Loop
End Sub

Sub SelectEntity(mtEntNode As IXMLDOMElement, swAsmDoc As IAssemblyDoc, swMath As IMathUtility)
Dim mtEntType As Integer
mtEntType = mtEntNode.selectSingleNode("type").Text
Dim SelectState As Boolean
SelectState = False
Dim SelectName As String
Dim ID As Object
Dim ComponentName As String
ComponentName = XMLImport.cpNameReplace(mtEntNode.getAttribute("component"))

Dim swSelMgr As ISelectionMgr
Set swSelMgr = swAsmDoc.SelectionManager
Dim swSelectData As ISelectData
Set swSelectData = swSelMgr.CreateSelectData()
swSelectData.Mark = 1

swSelMgr.SuspendSelectionList

Dim swRefCp As IComponent2
SelectName = ComponentName
SelectState = swAsmDoc.Extension.SelectByID2(SelectName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
Debug.Print SelectName
Debug.Print SelectState
If Not SelectState Then Exit Sub

Set swRefCp = swSelMgr.GetSelectedObject6(1, -1)

swSelMgr.ResumeSelectionList2 False

Dim swRefCp_p As IComponent2
Set swRefCp_p = swRefCp
Do Until swRefCp_p Is Nothing
    swRefCp_p.SetVisibility swComponentVisible, swThisConfiguration, ""
    Set swRefCp_p = swRefCp_p.GetParent()
Loop

Select Case mtEntType
    Case swSelEDGES, swSelFACES, swSelVERTICES
        swSelMgr.SuspendSelectionList
        
        Dim mtParamNodes As Variant
        Set mtParamNodes = mtEntNode.selectNodes("params/value")
        Dim j As Integer
        Dim mtParam(6) As Double
        For j = 0 To 6
            mtParam(j) = mtParamNodes(j).Text
        Next
    
        Dim nPt(2) As Double
        Dim vPt As Variant
        Dim mtEntPt As IMathPoint
        Dim mtEntVec As IMathVector
                
        ' pointX, Y, Z
        For j = 0 To 2
            nPt(j) = mtParam(j)
        Next
        vPt = nPt
        Set mtEntPt = swMath.CreatePoint((vPt))
        If Not swRefCp.Transform2 Is Nothing Then Set mtEntPt = mtEntPt.MultiplyTransform(swRefCp.Transform2)
        For j = 0 To 2
            mtParam(j) = mtEntPt.ArrayData(j)
        Next
        
        ' vectorI, J ,K
        For j = 0 To 2
            nPt(j) = mtParam(j + 3)
        Next
        vPt = nPt
        Set mtEntVec = swMath.CreateVector((vPt))
        If Not swRefCp.Transform2 Is Nothing Then Set mtEntVec = mtEntVec.MultiplyTransform(swRefCp.Transform2)
        Set mtEntVec = mtEntVec.Scale(-1)
        For j = 0 To 2
            mtParam(j + 3) = mtEntVec.ArrayData(j)
        Next
        
        'Set mtEntPt = mtEntPt.AddVector(mtEntVec.Scale(-0.0001)) 'ちょっと前に位置をずらして見つけやすくする
        
        
        mtParam(6) = mtParam(6) + 0.00001
        SelectState = swAsmDoc.MultiSelectByRay(mtParam, mtEntType, False)
        Debug.Print SelectState
        
        'Set GetEntity = swSelMgr.GetSelectedObject6(1, -1)
        
        Dim swEntity As Object
        Dim i As Integer
        For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
            Set swEntity = swSelMgr.GetSelectedObject6(i, -1)
            SelectState = True
            
            '属するBodyの中心を確認
            Set mtParamNodes = mtEntNode.selectNodes("body/value")
            Dim mtEntBodyParams As Variant
            If swMateEnt.ReferenceType2 = swSelEDGES Or swMateEnt.ReferenceType2 = swSelFACES Then
                mtEntBodyParams = swEntity.GetBody().GetMassProperties(1)
            ElseIf swMateEnt.ReferenceType2 = swSelVERTICES Then
                mtEntBodyParams = swEntity.GetEdges()(0).GetBody().GetMassProperties(1)
            End If
            For j = 0 To 2
                If Abs(CDbl(mtParamNodes(j).Text) - mtEntBodyParams(j)) < 0.00001 Then SelectState = False: Exit For
            Next
            
            'Edgeのパラメーターを確認
            If swMateEnt.ReferenceType2 = swSelEDGES Then
                Dim mtEdgeParams As Variant
                mtEdgeParams = swEntity.GetCurve().CircleParams
                If Not IsNull(mtEdgeParams) Then
                    Set mtParamNodes = mtEntNode.selectNodes("circleparams/value")
                    For j = 0 To 6
                        If Abs(CDbl(mtParamNodes(j).Text) - mtEdgeParams(j)) < 0.00001 Then SelectState = False: Exit For
                    Next
                End If
                
                mtEdgeParams = swEntity.GetCurve().LineParams
                If Not IsNull(mtEdgeParams) Then
                    Set mtParamNodes = mtEntNode.selectNodes("lineparams/value")
                    For j = 0 To 5
                        If Abs(CDbl(mtParamNodes(j).Text) - mtEdgeParams(j)) < 0.00001 Then SelectState = False: Exit For
                    Next
                End If
            End If
            
            '部品座標に戻す
            Set mtParamNodes = mtEntNode.selectNodes("params/value")
            For j = 0 To 6
                mtParam(j) = mtParamNodes(j).Text
            Next
            
            'Faceのパラメーターを出力
            If swMateEnt.ReferenceType2 = swSelFACES Then
                Set mtParamNodes = mtEntNode.selectNodes("faceparams/value")
                Dim mtFaceParams As Variant
                mtFaceParams = swEntity.GetSurface().EvaluateAtPoint(mtParam(0), mtParam(1), mtParam(2))
                If IsEmpty(mtFaceParams) Then
                    SelectState = False
                Else
                    For j = 0 To 10
                        If Abs(CDbl(mtParamNodes(j).Text) - mtFaceParams(j)) < 0.00001 Then SelectState = False: Exit For
                    Next
                End If
            End If
            
            'Vertexの位置確認
            If swMateEnt.ReferenceType2 = swSelVERTICES Then
                For j = 0 To 2
                    If Abs(CDbl(swEntity.GetPoint().ArrayData(j)) - mtParam(j)) < 0.00001 Then SelectState = False: Exit For
                Next
            End If
            
            
            If SelectState Then
                swSelMgr.ResumeSelectionList2 False
                SelectState = swSelMgr.AddSelectionListObject(swEntity, swSelectData)
                Debug.Print SelectState
                Exit For
            End If
        Next
        
        If Not SelectState Then
            swSelMgr.ResumeSelectionList2 False
        End If

    Case swSelDATUMPLANES, swSelDATUMAXES, swSelDATUMPOINTS
        
        SelectName = mtEntNode.selectSingleNode("name").Text
        If ComponentName <> "" Then SelectName = SelectName & "@" & ComponentName
        SelectState = swAsmDoc.Extension.SelectByID2(SelectName, GetSelTypeString(mtEntType), 0, 0, 0, True, 1, Nothing, 0)
        'Set GetEntity = swSelMgr.GetSelectedObject6(1, -1)
        Debug.Print SelectName
        Debug.Print SelectState
        
    Case swSelSKETCHSEGS, swSelSKETCHPOINTS, swSelEXTSKETCHSEGS, swSelEXTSKETCHPOINTS
        swSelMgr.SuspendSelectionList
        
        SelectName = mtEntNode.selectSingleNode("name").Text
        If ComponentName <> "" Then SelectName = SelectName & "@" & ComponentName
        Set ID = mtEntNode.selectNodes("id")
        SelectState = swAsmDoc.Extension.SelectByID2(SelectName, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Debug.Print SelectName
        Debug.Print SelectState
        If SelectState Then
            
            Dim swSketch As ISketch
            Set swSketch = swSelMgr.GetSelectedObject6(1, -1).GetSpecificFeature2()
            
            
            Dim swSketchEntities As Variant
            Select Case mtEntType
                Case swSelSKETCHSEGS, swSelEXTSKETCHSEGS
                    swSketchEntities = swSketch.GetSketchSegments()
                Case swSelSKETCHPOINTS, swSelEXTSKETCHPOINTS
                    swSketchEntities = swSketch.GetSketchPoints2()
            End Select

            swSelMgr.ResumeSelectionList2 False

            Dim swSketchEntitySingle As Variant
            For Each swSketchEntitySingle In swSketchEntities
                If CStr(swSketchEntitySingle.GetID()(0)) = ID(0).Text And CStr(swSketchEntitySingle.GetID()(1)) = ID(1).Text Then
                    SelectState = swSketchEntitySingle.Select4(True, swSelectData)
                    'Set GetEntity = swSelMgr.GetSelectedObject6(1, -1)
                    Exit For
                End If
            Next
        End If
        
    Case Else
        MsgBox "Unsupported type " & mtEntType
End Select

'If Not SelectState Then MsgBox "select failed"
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
