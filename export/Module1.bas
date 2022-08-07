Attribute VB_Name = "Module1"
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

Dim asmDoc As IAssemblyDoc
Set asmDoc = swModel

Dim DOMDoc As MSXML2.DOMDocument60
Set DOMDoc = New MSXML2.DOMDocument60
Call DOMDoc.appendChild(DOMDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8"""))

Dim RootNode As MSXML2.IXMLDOMNode
Set RootNode = DOMDoc.appendChild(DOMDoc.createNode(NODE_ELEMENT, "assembly", ""))

Dim cpArray As Variant
cpArray = asmDoc.GetComponents(True)
Dim i As Integer
For i = LBound(cpArray) To UBound(cpArray)
    Dim component As IComponent2
    Set component = cpArray(i)
    
    Dim cpNode As MSXML2.IXMLDOMNode
    Set cpNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Dim cpIdAttr As MSXML2.IXMLDOMAttribute
    Set cpIdAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "id", ""))
    cpIdAttr.nodeValue = component.GetID()


    Dim cpPathAttr As MSXML2.IXMLDOMAttribute
    Set cpPathAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "path", ""))
    cpPathAttr.nodeValue = component.GetPathName()

    
    Dim cpModel As ModelDoc2
    Set cpModel = component.GetModelDoc2()
    
    Dim cpTypeAttr As MSXML2.IXMLDOMAttribute
    Set cpTypeAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "type", ""))
    cpTypeAttr.nodeValue = cpModel.GetType()
    
Next
    

DOMDoc.Save swModel.GetPathName() + ".xml"
End Sub
