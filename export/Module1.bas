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

DOMDoc.Save swModel.GetPathName() + ".xml"
End Sub
