Attribute VB_Name = "XMLImport"
Dim swApp As Object
Sub main()
'Set swApp = Application.SldWorks
Set swApp = GetObject(, "Sldworks.Application")

Dim OpenFile As Variant
OpenFile = swApp.ActiveDoc.GetPathName() + ".xml"

Dim swTemplate As String
swTemplate = swApp.GetDocumentTemplate(swDocASSEMBLY, "", 0, 0, 0)
Dim swModel As ModelDoc2
Set swModel = swApp.NewDocument(swTemplate, 0, 0, 0)

Dim AsmDoc As IAssemblyDoc
Set AsmDoc = swModel

Dim DOMDoc As MSXML2.DOMDocument60
Set DOMDoc = New MSXML2.DOMDocument60
DOMDoc.Load OpenFile

Dim cpNode As MSXML2.IXMLDOMElement
For Each cpNode In DOMDoc.selectNodes("/assembly/components/component")
    Dim cpId As Integer
    Dim cpPath As String
    Dim cpType As Integer
    Dim cpConfiguration As String
    Dim cpTransformNodes As MSXML2.IXMLDOMNodeList
    cpId = cpNode.getAttribute("id")
    cpPath = cpNode.getAttribute("path")
    cpType = cpNode.selectSingleNode("type").Text
    cpConfiguration = cpNode.selectSingleNode("configuration").Text
    Set cpTransformNodes = cpNode.selectNodes("transform/value")

    Dim component As IComponent2
    Set component = AsmDoc.AddComponent5(cpPath, swAddComponentConfigOptions_CurrentSelectedConfig, "", True, cpConfiguration, 0, 0, 0)
    
    Dim j As Integer
    For j = 0 To 12
        component.Transform2.ArrayData(j) = cpTransformNodes(j).Text
    Next
    
Next
    
End Sub

