Attribute VB_Name = "XMLImport"
Dim swApp As Object

'xmlからアセンブリ生成
Sub main()
Set swApp = Application.SldWorks

Dim swMath As IMathUtility
Set swMath = swApp.GetMathUtility()

Dim OpenFile As Variant
OpenFile = swApp.ActiveDoc.GetPathName() + ".xml"

'新規アセンブリ
Dim swTemplate As String
swTemplate = swApp.GetDocumentTemplate(swDocASSEMBLY, "", 0, 0, 0)
Dim swModel As ModelDoc2
Set swModel = swApp.NewDocument(swTemplate, 0, 0, 0)

Dim swAsmDoc As IAssemblyDoc
Set swAsmDoc = swModel

Dim DOMDoc As DOMDocument60
Set DOMDoc = New DOMDocument60
DOMDoc.Load OpenFile

Dim cpIDReplacement As Object
Set cpIDReplacement = CreateObject("Scripting.Dictionary")


Dim cpNode As IXMLDOMElement

For Each cpNode In DOMDoc.selectNodes("/assembly/components/component")
    Dim cpID As String
    Dim cpPath As String
    Dim cpType As Integer
    Dim cpConfiguration As String
    Dim cpSolving As Integer
    Dim cpVisible As Boolean
    Dim cpSuppression As Integer
    
    cpID = cpNode.getAttribute("id")
    cpPath = cpNode.getAttribute("path")
    cpType = cpNode.selectSingleNode("type").Text
    cpConfiguration = cpNode.selectSingleNode("configuration").Text
    cpSolving = cpNode.selectSingleNode("solving").Text
    cpVisible = cpNode.selectSingleNode("visible").Text
    cpSuppression = cpNode.selectSingleNode("suppression").Text
    
    Dim swComponent As IComponent2
    Set swComponent = swAsmDoc.AddComponent5(cpPath, swAddComponentConfigOptions_CurrentSelectedConfig, "", True, cpConfiguration, 0, 0, 0)
    
    cpIDReplacement.Add cpID, swComponent.GetID()
    cpID = swComponent.GetID()
    
    swModel.Extension.SelectByID2 swComponent.GetSelectByIDString(), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0
    swAsmDoc.CompConfigProperties6 cpSuppression, cpSolving, cpVisible, False, "", False, False, False
    
    ApplyComponentProps cpNode, swComponent, swMath, cpIDReplacement, cpID
    
Next


Dim mtNode As IXMLDOMElement
For Each mtNode In DOMDoc.selectNodes("/assembly/mates/mate")
    Dim mtType As Integer
    Dim mtAlignment As Integer
    
    mtType = mtNode.selectSingleNode("type").Text
    mtAlignment = mtNode.selectSingleNode("alignment").Text
    
    swModel.ClearSelection2 True
    
    Dim mtEntNode As IXMLDOMElement
    For Each mtEntNode In mtNode.selectNodes("entity")
        Dim mtRefCpID As Variant
        Dim mtRefCpIDPart As String
        mtRefCpIDPart = ""
        Dim mtEntType As Integer
        Dim j As Integer
        
        mtRefCpID = Split(mtEntNode.getAttribute("component-id"), "/")
        For j = 0 To UBound(mtRefCpID)
            Dim replacedID As String
            replacedID = mtRefCpID(j)
            
            If j > 0 Then mtRefCpIDPart = mtRefCpIDPart & "/"
            If cpIDReplacement.Exists(mtRefCpIDPart & replacedID) Then replacedID = cpIDReplacement(mtRefCpIDPart & replacedID)
            
            mtRefCpIDPart = mtRefCpIDPart & replacedID
        Next
        mtRefCpID = Split(mtRefCpIDPart, "/")
        
        mtEntType = mtEntNode.selectSingleNode("type").Text
        
        Set mtParamNodes = mtEntNode.selectNodes("params/value")
        
        Dim mtParam(7) As Double
        For j = 0 To 7
            mtParam(j) = mtParamNodes(j).Text
        Next

        Dim swEntCp As IComponent2
        Set swEntCp = swAsmDoc.GetComponentByID(mtRefCpID(0))
        For j = 1 To UBound(mtRefCpID)
            Set swEntCp = swEntCp.GetModelDoc2().GetComponentByID(mtRefCpID(j))
        Next
        
        Dim swEntModel As IModelDoc2
        Set swEntModel = swEntCp.GetModelDoc2()
        
        Dim SelectState As Boolean
        'SelectState = swEntModel.Extension.SelectByID2("", SelType.GetSelTypeString(mtEntType), mtParam(0), mtParam(1), mtParam(2), True, 1, Nothing, 0)
        SelectState = swEntModel.Extension.SelectByRay(mtParam(0), mtParam(1), mtParam(2), mtParam(3), mtParam(4), mtParam(5), 0.001, mtEntType, True, 1, 0)
        MsgBox SelectState
    Next
    
    Dim mtData As IMateFeatureData
    Set mtData = swAsmDoc.CreateMateData(mtType)
    If mtType = swMateCOINCIDENT Then
        Dim mtDataCasted As ICoincidentMateFeatureData
        Set mtDataCasted = mtData
        
        mtDataCasted.MateAlignment = mtAlignment
    End If
    swAsmDoc.CreateMate mtDataCasted
Next

End Sub

Sub ApplyComponentProps(cpNode As IXMLDOMElement, swComponent As IComponent2, swMath As IMathUtility, cpIDReplacement As Object, cpID As String)

Dim cpTransformNodes As IXMLDOMNodeList
Dim cpChildren As IXMLDOMNodeList

Set cpTransformNodes = cpNode.selectNodes("transform/value")
Set cpChildren = cpNode.selectNodes("components/component")

Dim TransformArray(15) As Double
Dim j As Integer
For j = 0 To 12
    TransformArray(j) = cpTransformNodes(j).Text
Next
swComponent.Transform2 = swMath.CreateTransform(TransformArray)

Dim swChild As IComponent2
Dim swChildren As Variant
swChildren = swComponent.GetChildren()
Dim i As Integer
For i = LBound(swChildren) To UBound(swChildren)
    Set swChild = swChildren(i)
    Dim cpElement As IXMLDOMElement
    For Each cpElement In cpChildren
        If cpElement.getAttribute("id") = swChild.GetID() Then
            ApplyComponentProps cpElement, swChild, swMath, cpIDReplacement, cpID
            Exit For
        End If
    Next
Next
    
End Sub
