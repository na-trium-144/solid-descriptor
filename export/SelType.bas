Attribute VB_Name = "SelType"
Function ExportEntity(swMateEnt As IMateEntity2, mtEntNode As IXMLDOMNode, DOMDoc As DOMDocument60, swMath As IMathUtility) As Boolean
Dim mtSubNode As IXMLDOMNode
Dim mtAttr As IXMLDOMAttribute

Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "type", ""))
mtSubNode.Text = swMateEnt.ReferenceType2

' component.Name2
Set mtAttr = mtEntNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "component", ""))
'mtAttr.nodeValue = swRefCp.Name2
mtAttr.nodeValue = swMateEnt.ReferenceComponent.GetSelectByIDString()

Dim nPt(2) As Double
Dim vPt As Variant
Dim mtEntPt As IMathPoint
Dim ClosestPt As Variant
Dim mtParams As Variant
Dim j As Integer
Dim mtParamValueNode As IXMLDOMNode

If swMateEnt.Reference Is Nothing Then
    Debug.Print swMateEnt.ReferenceComponent.GetSelectByIDString() & ", Reference is Nothing"
    ExportEntity = False
    Exit Function
End If

Select Case swMateEnt.ReferenceType2
    Case swSelEDGES
        '属するBodyを出力
        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "body", ""))
        mtSubNode.Text = swMateEnt.Reference.GetBody().Name
        
        'Edgeのパラメーターを出力
        
        mtParams = swMateEnt.Reference.GetCurve().CircleParams
        If Not IsNull(mtParams) Then
            Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "circleparams", ""))
            For j = 0 To 6
                Set mtParamValueNode = mtSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
                mtParamValueNode.Text = mtParams(j)
            Next
        End If
        
        mtParams = swMateEnt.Reference.GetCurve().LineParams
        If Not IsNull(mtParams) Then
            Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "lineparams", ""))
            For j = 0 To 5
                Set mtParamValueNode = mtSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
                mtParamValueNode.Text = mtParams(j)
            Next
        End If
    
    Case swSelFACES
        
        For j = 0 To 2
            nPt(j) = swMateEnt.EntityParams(j)
        Next
        vPt = nPt
        Set mtEntPt = swMath.CreatePoint((vPt))
        If Not swMateEnt.ReferenceComponent.Transform2 Is Nothing Then Set mtEntPt = mtEntPt.MultiplyTransform(swMateEnt.ReferenceComponent.Transform2.Inverse())

        'EntityParamsの点が面上にないことがあり、その場合は面上で一番近い点にする
        ClosestPt = swMateEnt.Reference.GetClosestPointOn(mtEntPt.ArrayData(0), mtEntPt.ArrayData(1), mtEntPt.ArrayData(2))
        For j = 0 To 2
            nPt(j) = ClosestPt(j)
        Next
        vPt = nPt
        Set mtEntPt = swMath.CreatePoint((vPt))
        
        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "point", ""))
        For j = 0 To 2
            Set mtParamValueNode = mtSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
            mtParamValueNode.Text = mtEntPt.ArrayData(j)
        Next
        
        '属するBodyを出力
        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "body", ""))
        mtSubNode.Text = swMateEnt.Reference.GetBody().Name
        
        'Faceのパラメーターを出力
        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "faceparams", ""))
        mtParams = swMateEnt.Reference.GetSurface().EvaluateAtPoint(mtEntPt.ArrayData(0), mtEntPt.ArrayData(1), mtEntPt.ArrayData(2))
        For j = 0 To 10
            Set mtParamValueNode = mtSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
            mtParamValueNode.Text = mtParams(j)
        Next

    
    Case swSelVERTICES
        
        Set mtEntPt = swMateEnt.Reference.GetPoint()

        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "point", ""))
        For j = 0 To 2
            Set mtParamValueNode = mtSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
            mtParamValueNode.Text = mtEntPt.ArrayData(j)
        Next
        
                
        '属するBodyを出力
        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "body", ""))
        mtSubNode.Text = swMateEnt.Reference.GetBody().Name

    Case swSelDATUMPLANES, swSelDATUMAXES, swSelDATUMPOINTS
        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "name", ""))
        mtSubNode.Text = swMateEnt.Reference.Name
        
    Case swSelSKETCHSEGS, swSelSKETCHPOINTS, swSelEXTSKETCHSEGS, swSelEXTSKETCHPOINTS
        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "name", ""))
        mtSubNode.Text = swMateEnt.Reference.GetSketch().Name
        
        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "id", ""))
        mtSubNode.Text = swMateEnt.Reference.GetID()(0)
        Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "id", ""))
        mtSubNode.Text = swMateEnt.Reference.GetID()(1)
        
    Case Else
        MsgBox "Unsupported type " & swMateEnt.ReferenceType2
End Select

ExportEntity = True

End Function

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
SelectState = swAsmDoc.Extension.SelectByID2(SelectName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
Debug.Print SelectName
Debug.Print SelectState
If Not SelectState Then swSelMgr.ResumeSelectionList2 False: Exit Sub

Set swRefCp = swSelMgr.GetSelectedObject6(1, -1)

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

Dim SelectedObjCountBefore As Integer
SelectedObjCountBefore = swSelMgr.GetSelectedObjectCount2(1)

swSelMgr.SuspendSelectionList

Dim swRefCp As IComponent2
SelectName = ComponentName
SelectState = swAsmDoc.Extension.SelectByID2(SelectName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
Debug.Print SelectName
Debug.Print SelectState
If Not SelectState Then
    If SelectedObjCountBefore > 0 Then swSelMgr.ResumeSelectionList2 False
    Exit Sub
End If

Set swRefCp = swSelMgr.GetSelectedObject6(1, -1)

If SelectedObjCountBefore > 0 Then swSelMgr.ResumeSelectionList2 False

'Dim swRefCp_p As IComponent2
'Set swRefCp_p = swRefCp
'Do Until swRefCp_p Is Nothing
'    swRefCp_p.SetVisibility swComponentVisible, swThisConfiguration, ""
'    Set swRefCp_p = swRefCp_p.GetParent()
'Loop
        
Dim swBodies As Variant
Dim swBody As Variant

Dim nPt(2) As Double
Dim vPt As Variant
Dim mtEntPt As IMathPoint
Dim mtParamNodes As Variant
Dim mtParams As Variant

Dim mtBodyName As String

Dim swEntity As Variant
Dim i As Integer
Dim j As Integer

swSelMgr.SuspendSelectionList
' Appendなしでselectし、最後にResumeAppendする

Select Case mtEntType
    Case swSelEDGES
        
        mtBodyName = mtEntNode.selectSingleNode("body").Text
        
        swBodies = swRefCp.GetBodies3(swAllBodies, swNormalBody_e)
        If IsEmpty(swBodies) Then
            If SelectedObjCountBefore > 0 Then swSelMgr.ResumeSelectionList2 False
            Exit Sub
        End If
        
        For i = 0 To UBound(swBodies)
            Set swBody = swBodies(i)

            For Each swEntity In swBody.GetEdges()
            
                SelectState = True

                'Edgeのパラメーターを確認
                mtParams = swEntity.GetCurve().CircleParams
                If Not IsNull(mtParams) Then
                    Set mtParamNodes = mtEntNode.selectNodes("circleparams/value")
                    If mtParamNodes.Length = 0 Then
                        SelectState = False
                    Else
                        For j = 0 To 6
                            If Abs(CDbl(mtParamNodes(j).Text) - mtParams(j)) > 0.00001 Then SelectState = False: Exit For
                        Next
                    End If
                End If
                
                mtParams = swEntity.GetCurve().LineParams
                If Not IsNull(mtParams) Then
                    Set mtParamNodes = mtEntNode.selectNodes("lineparams/value")
                    If mtParamNodes.Length = 0 Then
                        SelectState = False
                    Else
                        For j = 0 To 5
                            If Abs(CDbl(mtParamNodes(j).Text) - mtParams(j)) > 0.00001 Then SelectState = False: Exit For
                        Next
                    End If
                End If
                
                If SelectState Then
                    swAsmDoc.ClearSelection2 True
                    SelectState = swSelMgr.AddSelectionListObject(swEntity, swSelectData)
                    Debug.Print SelectState
                    Exit For
                End If
                
            Next
                        
            If SelectState Then
                If swBody.Name = mtBodyName Or UBound(swBodies) = 0 Then Exit For
                ' else まだ探す
            End If
        Next

    Case swSelFACES
        
        mtBodyName = mtEntNode.selectSingleNode("body").Text
        
        Set mtParamNodes = mtEntNode.selectNodes("point/value")
        For j = 0 To 2
            nPt(j) = mtParamNodes(j).Text
        Next
        
        swBodies = swRefCp.GetBodies3(swAllBodies, swNormalBody_e)
        If IsEmpty(swBodies) Then
            If SelectedObjCountBefore > 0 Then swSelMgr.ResumeSelectionList2 False
            Exit Sub
        End If
        
        For i = 0 To UBound(swBodies)
            Set swBody = swBodies(i)
            
            For Each swEntity In swBody.GetFaces()
            
                SelectState = True

                'Faceのパラメーター
                Set mtParamNodes = mtEntNode.selectNodes("faceparams/value")
                mtParams = swEntity.GetSurface().EvaluateAtPoint(nPt(0), nPt(1), nPt(2))
                If IsEmpty(mtParams) Then
                    SelectState = False
                Else
                    For j = 0 To 10
                        If Abs(CDbl(mtParamNodes(j).Text) - mtParams(j)) > 0.00001 Then SelectState = False: Exit For
                    Next
                End If
                
                If SelectState Then
                    swAsmDoc.ClearSelection2 True
                    SelectState = swSelMgr.AddSelectionListObject(swEntity, swSelectData)
                    Debug.Print SelectState
                    Exit For
                End If
                
            Next
                        
            If SelectState Then
                If swBody.Name = mtBodyName Or UBound(swBodies) = 0 Then Exit For
                ' else まだ探す
            End If
        Next


    Case swSelVERTICES
        
        mtBodyName = mtEntNode.selectSingleNode("body").Text
        
        Set mtParamNodes = mtEntNode.selectNodes("point/value")
        For j = 0 To 2
            nPt(j) = mtParamNodes(j).Text
        Next
        
        
        swBodies = swRefCp.GetBodies3(swAllBodies, swNormalBody_e)
        If IsEmpty(swBodies) Then
            If SelectedObjCountBefore > 0 Then swSelMgr.ResumeSelectionList2 False
            Exit Sub
        End If
        
        For i = 0 To UBound(swBodies)
            Set swBody = swBodies(i)
            
            For Each swEntity In swBody.GetVertices()
            
                SelectState = True

                'Vertexの位置確認
                For j = 0 To 2
                    If Abs(CDbl(swEntity.GetPoint().ArrayData(j)) - nPt(j)) > 0.00001 Then SelectState = False: Exit For
                Next

                If SelectState Then
                    swAsmDoc.ClearSelection2 True
                    SelectState = swSelMgr.AddSelectionListObject(swEntity, swSelectData)
                    Debug.Print SelectState
                    Exit For
                End If
                
            Next
                        
            If SelectState Then
                If swBody.Name = mtBodyName Or UBound(swBodies) = 0 Then Exit For
                ' else まだ探す
            End If
        Next


    Case swSelDATUMPLANES, swSelDATUMAXES, swSelDATUMPOINTS
        
        SelectName = mtEntNode.selectSingleNode("name").Text
        If ComponentName <> "" Then SelectName = SelectName & "@" & ComponentName
        SelectState = swAsmDoc.Extension.SelectByID2(SelectName, GetSelTypeString(mtEntType), 0, 0, 0, False, 1, Nothing, 0)
        'Set GetEntity = swSelMgr.GetSelectedObject6(1, -1)
        Debug.Print SelectName
        Debug.Print SelectState
        
    Case swSelSKETCHSEGS, swSelSKETCHPOINTS, swSelEXTSKETCHSEGS, swSelEXTSKETCHPOINTS
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

            Dim swSketchEntitySingle As Variant
            For Each swSketchEntitySingle In swSketchEntities
                If CStr(swSketchEntitySingle.GetID()(0)) = ID(0).Text And CStr(swSketchEntitySingle.GetID()(1)) = ID(1).Text Then
                    SelectState = swSketchEntitySingle.Select4(False, swSelectData)
                    'Set GetEntity = swSelMgr.GetSelectedObject6(1, -1)
                    Exit For
                End If
            Next
        End If
        
    Case Else
        MsgBox "Unsupported type " & mtEntType
End Select

'If Not SelectState Then MsgBox "select failed"

If SelectedObjCountBefore > 0 Then swSelMgr.ResumeSelectionList2 True

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
