Attribute VB_Name = "XMLImport"
Dim swApp As Object
Sub main()
'Set swApp = Application.SldWorks
Set swApp = GetObject(, "Sldworks.Application")

Dim swMath As IMathUtility
Set swMath = swApp.GetMathUtility()

Dim OpenFile As Variant
OpenFile = swApp.ActiveDoc.GetPathName() + ".xml"

Dim swTemplate As String
swTemplate = swApp.GetDocumentTemplate(swDocASSEMBLY, "", 0, 0, 0)
Dim swModel As ModelDoc2
Set swModel = swApp.NewDocument(swTemplate, 0, 0, 0)

Dim AsmDoc As IAssemblyDoc
Set AsmDoc = swModel

Dim DOMDoc As DOMDocument60
Set DOMDoc = New DOMDocument60
DOMDoc.Load OpenFile

Dim xmlcpNode As IXMLDOMElement
    
For Each xmlcpNode In DOMDoc.selectNodes("/assembly/components/component")
    Dim xmlcpId As Integer
    Dim xmlcpPath As String
    Dim xmlcpType As Integer
    Dim xmlcpConfiguration As String
    Dim xmlcpSolving As Integer
    Dim xmlcpVisible As Boolean
    Dim xmlcpSuppression As Integer
    
    xmlcpId = xmlcpNode.getAttribute("id")
    xmlcpPath = xmlcpNode.getAttribute("path")
    xmlcpType = xmlcpNode.selectSingleNode("type").Text
    xmlcpConfiguration = xmlcpNode.selectSingleNode("configuration").Text
    xmlcpSolving = xmlcpNode.selectSingleNode("solving").Text
    xmlcpVisible = xmlcpNode.selectSingleNode("visible").Text
    xmlcpSuppression = xmlcpNode.selectSingleNode("suppression").Text
    
    Dim component As IComponent2
    Set component = AsmDoc.AddComponent5(xmlcpPath, swAddComponentConfigOptions_CurrentSelectedConfig, "", True, xmlcpConfiguration, 0, 0, 0)
    
    swModel.Extension.SelectByID2 component.GetSelectByIDString(), "COMPONENT", 0, 0, 0, False, 0, Nothing, 0
    AsmDoc.CompConfigProperties6 xmlcpSuppression, xmlcpSolving, xmlcpVisible, False, "", False, False, False
    
    ApplyComponentProps xmlcpNode, component, swMath
    
Next
End Sub

Sub ApplyComponentProps(xmlcpNode As IXMLDOMElement, component As IComponent2, swMath As IMathUtility)


Dim xmlcpTransformNodes As IXMLDOMNodeList
Dim xmlcpChildren As IXMLDOMNodeList

Set xmlcpTransformNodes = xmlcpNode.selectNodes("transform/value")
Set xmlcpChildren = xmlcpNode.selectNodes("components/component")

Dim TransformArray(15) As Double
Dim j As Integer
For j = 0 To 12
    TransformArray(j) = xmlcpTransformNodes(j).Text
Next
component.Transform2 = swMath.CreateTransform(TransformArray)

Dim cpChild As IComponent2
Dim cpChildren As Variant
cpChildren = component.GetChildren()
Dim i As Integer
For i = LBound(cpChildren) To UBound(cpChildren)
    Set cpChild = cpChildren(i)
    Dim xmlcpElement As IXMLDOMElement
    For Each xmlcpElement In xmlcpChildren
        If xmlcpElement.getAttribute("id") = cpChild.GetID() Then
            ApplyComponentProps xmlcpElement, cpChild, swMath
            Exit For
        End If
    Next
Next
    
End Sub
