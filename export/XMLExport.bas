Attribute VB_Name = "XMLExport"
Dim swApp As Object
Sub main()
'Set swApp = Application.SldWorks
Set swApp = GetObject(, "Sldworks.Application")

Dim swModel As ModelDoc2
Set swModel = swApp.ActiveDoc
If swModel.GetType() <> swDocASSEMBLY Then
    MsgBox "Not an assembly"
    Exit Sub
End If

Dim AsmDoc As IAssemblyDoc
Set AsmDoc = swModel

Dim DOMDoc As MSXML2.DOMDocument60
Set DOMDoc = New MSXML2.DOMDocument60
DOMDoc.appendChild DOMDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")

Dim RootNode As MSXML2.IXMLDOMNode
Set RootNode = DOMDoc.appendChild(DOMDoc.createNode(NODE_ELEMENT, "assembly", ""))

Dim ComponentsNode As MSXML2.IXMLDOMNode
Set ComponentsNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "components", ""))

Dim cpArray As Variant
cpArray = AsmDoc.GetComponents(True)
Dim i As Integer
For i = LBound(cpArray) To UBound(cpArray)
    Dim component As IComponent2
    Set component = cpArray(i)
    Dim cpModel As ModelDoc2
    Set cpModel = component.GetModelDoc2()
    
    Dim cpNode As MSXML2.IXMLDOMNode
    Set cpNode = ComponentsNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Dim cpAttr As MSXML2.IXMLDOMAttribute
    Dim cpSubNode As MSXML2.IXMLDOMNode
    
    Set cpAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "id", ""))
    cpAttr.nodeValue = component.GetID()

    Set cpAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "path", ""))
    cpAttr.nodeValue = component.GetPathName()

    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "type", ""))
    cpSubNode.Text = cpModel.GetType()
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "configuration", ""))
    cpSubNode.Text = component.ReferencedConfiguration
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "transform", ""))
    
    Dim j As Integer
    For j = 0 To 12
        Dim cpTransformValueNode As MSXML2.IXMLDOMNode
        Set cpTransformValueNode = cpSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
        cpTransformValueNode.Text = component.Transform2.ArrayData(j)
    Next
    
Next
    
DOMDoc.loadXML Indent.Indent(DOMDoc.xml)
DOMDoc.Save swModel.GetPathName() + ".xml"
End Sub
