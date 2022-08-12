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

Dim DOMDoc As DOMDocument60
Set DOMDoc = New DOMDocument60
DOMDoc.appendChild DOMDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")

Dim RootNode As IXMLDOMNode
Set RootNode = DOMDoc.appendChild(DOMDoc.createNode(NODE_ELEMENT, "assembly", ""))

Dim ComponentsNode As IXMLDOMNode
Set ComponentsNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "components", ""))

Dim cpArray As Variant
cpArray = AsmDoc.GetComponents(True)
Dim i As Integer
For i = LBound(cpArray) To UBound(cpArray)
    Dim component As IComponent2
    Set component = cpArray(i)
    'Dim cpModel As ModelDoc2
    'Set cpModel = component.GetModelDoc2()
    Dim cpPath As String
    cpPath = component.GetPathName()
    Dim cpExtension As String
    cpExtension = UCase(Right(cpPath, 7))
    
    Dim cpNode As IXMLDOMNode
    Set cpNode = ComponentsNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Dim cpAttr As IXMLDOMAttribute
    Dim cpSubNode As IXMLDOMNode
    
    Set cpAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "id", ""))
    cpAttr.nodeValue = component.GetID()

    Set cpAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "path", ""))
    cpAttr.nodeValue = cpPath

    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "type", ""))
    'cpSubNode.Text = cpModel.GetType()
    If cpExtension = ".SLDASM" Then
        cpSubNode.Text = swDocASSEMBLY
    ElseIf cpExtension = ".SLDPRT" Then
        cpSubNode.Text = swDocPART
    End If
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "configuration", ""))
    cpSubNode.Text = component.ReferencedConfiguration
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "solving", ""))
    cpSubNode.Text = component.Solving
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "suppression", ""))
    cpSubNode.Text = component.GetSuppression2()
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "visible", ""))
    cpSubNode.Text = component.Visible
    
    ExportComponentProps DOMDoc, cpNode, component
    
    
Next


DOMDoc.loadXML Indent.Indent(DOMDoc.xml)
DOMDoc.Save swModel.GetPathName() + ".xml"
End Sub

Sub ExportComponentProps(DOMDoc As DOMDocument60, cpNode As IXMLDOMNode, component As IComponent2)
Dim cpAttr As IXMLDOMAttribute
Dim cpSubNode As IXMLDOMNode
    
Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "transform", ""))

Dim j As Integer
For j = 0 To 15
    Dim cpTransformValueNode As IXMLDOMNode
    Set cpTransformValueNode = cpSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
    cpTransformValueNode.Text = component.Transform2.ArrayData(j)
Next

Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "components", ""))

Dim cpChildren As Variant
cpChildren = component.GetChildren()

Dim i As Integer
For i = LBound(cpChildren) To UBound(cpChildren)
    Dim child As IComponent2
    Set child = cpChildren(i)
    
    Dim cpChildNode As IXMLDOMNode
    Set cpChildNode = cpSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Set cpAttr = cpChildNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "id", ""))
    cpAttr.nodeValue = child.GetID()
    
    ExportComponentProps DOMDoc, cpChildNode, child
Next

End Sub
