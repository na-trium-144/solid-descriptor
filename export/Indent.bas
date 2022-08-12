Attribute VB_Name = "Indent"
' https://www.depthbomb.net/?p=6917
Function Indent(ByVal xml As String) As String
Dim writer As MSXML2.MXXMLWriter60
Dim reader As MSXML2.SAXXMLReader60
Dim dom As MSXML2.DOMDocument60
Dim n As MSXML2.IXMLDOMNode
    Set writer = New MSXML2.MXXMLWriter60
    writer.omitXMLDeclaration = True
    writer.Indent = True
    
    Set reader = New MSXML2.SAXXMLReader60
    Set reader.contentHandler = writer
    reader.Parse xml
    Set dom = New MSXML2.DOMDocument60
    dom.loadXML xml
    Set n = dom.childNodes(0)
    dom.loadXML writer.output
    If n.nodeName = "xml" And n.NodeType = NODE_PROCESSING_INSTRUCTION Then
        dom.InsertBefore n, dom.childNodes(0)
    End If
    Indent = dom.xml
End Function
