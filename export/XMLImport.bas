Attribute VB_Name = "XMLImport"
Dim swApp As Object

Dim swMath As IMathUtility

Dim cpNameReplaceList As Object

Function cpNameReplace(cpName As String) As String
Dim oldName As Variant
For Each oldName In cpNameReplaceList
    cpName = Replace(cpName, oldName, cpNameReplaceList(oldName))
Next
cpNameReplace = cpName
End Function

'xmlからアセンブリ生成
Sub main()
Set swApp = Application.SldWorks
Set swMath = swApp.GetMathUtility()

Set cpNameReplaceList = CreateObject("Scripting.Dictionary")

Dim OpenFile As Variant
OpenFile = swApp.ActiveDoc.GetPathName() + ".xml"

'新規アセンブリ
Dim swTemplate As String
swTemplate = swApp.GetDocumentTemplate(swDocASSEMBLY, "", 0, 0, 0)
Dim swModel As ModelDoc2
Set swModel = swApp.NewDocument(swTemplate, 0, 0, 0)

Dim swAsmDoc As IAssemblyDoc
Set swAsmDoc = swModel

Dim swSelMgr As ISelectionMgr
Set swSelMgr = swModel.SelectionManager

Dim swConfMgr As IConfigurationManager
Set swConfMgr = swModel.ConfigurationManager

Dim DOMDoc As DOMDocument60
Set DOMDoc = New DOMDocument60
DOMDoc.Load OpenFile

Dim RootCpName As String
RootCpName = DOMDoc.selectSingleNode("/assembly/@name").Text

'Componentの作成

Dim cfNode As IXMLDOMElement
For Each cfNode In DOMDoc.selectNodes("/assembly/configurations/configuration")
    Dim ConfName As String
    ConfName = cfNode.getAttribute("name")
    
    Dim swConf As IConfiguration
    Set swConf = swModel.GetConfigurationByName(ConfName)
    If swConf Is Nothing Then Set swConf = swModel.AddConfiguration3(ConfName, "", "", 0)

    swModel.ShowConfiguration2 ConfName
    

    Dim cpNode As IXMLDOMElement
    For Each cpNode In DOMDoc.selectNodes("/assembly/toplevel/component")
        Dim cpName As String
        Dim cpPath As String
    
        cpName = cpNode.getAttribute("name")
        cpPath = cpNode.selectSingleNode("path").Text
        
        Dim swComponent As IComponent2
        Dim SelectState As Boolean
        SelectState = swAsmDoc.Extension.SelectByID2(cpNameReplace(cpName), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        If SelectState Then
            Set swComponent = swSelMgr.GetSelectedObject6(1, -1)
        Else
            Set swComponent = swAsmDoc.AddComponent5(cpPath, swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0, 0, 0)
            cpNameReplaceList.Add cpName, swComponent.GetSelectByIDString()
            swComponent.Select4 False, Nothing, False
        End If
        Debug.Print swComponent.GetSelectByIDString()
        
        Dim cpReference As String
        cpReference = cpNode.selectSingleNode("reference[@configuration=""" & ConfName & """]").Text
        
        swAsmDoc.CompConfigProperties6 swComponentFullyResolved, swComponentFlexibleSolving, True, False, cpReference, False, False, False

    Next
    
Next

Dim swChildCp As IComponent2
Dim cpElement As IXMLDOMElement
For Each cpElement In DOMDoc.selectNodes("/assembly/transform/component")
    SelectState = swAsmDoc.Extension.SelectByID2(cpNameReplace(cpElement.getAttribute("name")), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    Set swChildCp = swSelMgr.GetSelectedObject6(1, -1)
    
    Dim cpTransformNodes As IXMLDOMNodeList
    Set cpTransformNodes = cpElement.selectNodes("value")
    
    Dim TransformArray(15) As Double
    Dim j As Integer
    For j = 0 To 12
        TransformArray(j) = cpTransformNodes(j).Text
    Next
    swChildCp.Transform2 = swMath.CreateTransform(TransformArray)

Next



For Each cfNode In DOMDoc.selectNodes("/assembly/configurations/configuration")
    ConfName = cfNode.getAttribute("name")
    
    Set swConf = swModel.GetConfigurationByName(ConfName)
    swModel.ShowConfiguration2 ConfName

    For Each cpNode In DOMDoc.selectNodes("/assembly/toplevel/component")
        cpName = cpNode.getAttribute("name")
        
        SelectState = swAsmDoc.Extension.SelectByID2(cpNameReplace(cpName), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        Set swComponent = swSelMgr.GetSelectedObject6(1, -1)
        
        Dim cpSolving As Integer
        Dim cpSuppression As Integer
        Dim cpFixed As Boolean
        
        
        cpSolving = cpNode.selectSingleNode("solving[@configuration=""" & ConfName & """]").Text
        cpReference = cpNode.selectSingleNode("reference[@configuration=""" & ConfName & """]").Text
        cpSuppression = cpNode.selectSingleNode("suppression[@configuration=""" & ConfName & """]").Text
        cpFixed = Not cpNode.selectSingleNode("fixed[@configuration=""" & ConfName & """]") Is Nothing
        
        swAsmDoc.CompConfigProperties6 cpSuppression, cpSolving, True, False, cpReference, False, False, False

        If cpFixed Then
            swAsmDoc.FixComponent
        Else
            swAsmDoc.UnfixComponent
        End If

    Next
    
    
    Dim mtNode As IXMLDOMElement
    For Each mtNode In DOMDoc.selectNodes("/assembly/mates/mate")
        Dim mtName As String
        mtName = mtNode.getAttribute("name")
        
        If mtNode.selectSingleNode("active[@configuration=""" & ConfName & """]") Is Nothing Then
            SelectState = swAsmDoc.Extension.SelectByID2(cpNameReplace(mtName), "MATE", 0, 0, 0, False, 0, Nothing, 0)
            If SelectState Then
                swModel.EditSuppress2
            End If
        Else
            SelectState = swAsmDoc.Extension.SelectByID2(cpNameReplace(mtName), "MATE", 0, 0, 0, False, 0, Nothing, 0)
            If SelectState Then
                swModel.EditUnsuppress2
            Else
                
                swModel.ClearSelection2 True
            
                Dim mtEntityNodes As Object
                Set mtEntityNodes = mtNode.selectNodes("entity")
                
                Dim mtEntNode As IXMLDOMElement
                Dim i As Integer
                'For Each mtEntNode In mtNode.selectNodes("entity")
                For i = 0 To mtEntityNodes.Length - 1
                    Set mtEntNode = mtEntityNodes(i)
                    SelType.SelectEntity mtEntNode, swAsmDoc, swMath
                Next
                Debug.Print swSelMgr.GetSelectedObjectCount2(1)
    
                Dim mtType As Integer
                mtType = mtNode.selectSingleNode("type").Text
                
                Dim mtData As IMateFeatureData
                Set mtData = swAsmDoc.CreateMateData(mtType)
                
                If mtType = swMateCOINCIDENT Then
                    Dim mtDataCasted As ICoincidentMateFeatureData
                    Set mtDataCasted = mtData
                
                    mtDataCasted.MateAlignment = mtNode.selectSingleNode("alignment").Text
                End If
                
                Dim swMate As Object
                Set swMate = swAsmDoc.CreateMate(mtDataCasted)
                If Not swMate Is Nothing Then cpNameReplaceList.Add mtName, swMate.Name
            End If
        End If
    Next
Next

End Sub


Sub HideAllComponent(swComponent As IComponent2)
Dim cpChildren As IXMLDOMNodeList

swComponent.SetVisibility swComponentHidden, swThisConfiguration, ""
If swComponent.GetSuppression2() = swComponentSuppressed Then swComponent.SetSuppression2 swComponentResolved

Dim swChild As Variant
Dim swChildren As Variant
swChildren = swComponent.GetChildren()
For Each swChild In swChildren
    Dim swChildCp As IComponent2
    Set swChildCp = swChild
    HideAllComponent swChildCp
Next
    
End Sub

