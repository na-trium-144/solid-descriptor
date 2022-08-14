Attribute VB_Name = "XMLExport"
Dim swApp As Object
Sub main()
'Set swApp = Application.SldWorks
Set swApp = GetObject(, "Sldworks.Application")

Dim swMath As IMathUtility
Set swMath = swApp.GetMathUtility()

Dim swModel As ModelDoc2
Set swModel = swApp.ActiveDoc
If swModel.GetType() <> swDocASSEMBLY Then
    MsgBox "Not an assembly"
    Exit Sub
End If

Dim swAsmDoc As IAssemblyDoc
Set swAsmDoc = swModel

Dim DOMDoc As DOMDocument60
Set DOMDoc = New DOMDocument60
DOMDoc.appendChild DOMDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")

Dim RootNode As IXMLDOMNode
Set RootNode = DOMDoc.appendChild(DOMDoc.createNode(NODE_ELEMENT, "assembly", ""))

Dim ComponentsNode As IXMLDOMNode
Set ComponentsNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "components", ""))

Dim MatesNode As IXMLDOMNode
Set MatesNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "mates", ""))

Dim swCpArray As Variant
swCpArray = swAsmDoc.GetComponents(True)
Dim i As Integer
For i = LBound(swCpArray) To UBound(swCpArray)
    Dim swComponent As IComponent2
    Set swComponent = swCpArray(i)
    'Dim swCModel As ModelDoc2
    'Set swCpModel = swComponent.GetModelDoc2()
    Dim cpPath As String
    cpPath = swComponent.GetPathName()
    Dim cpExtension As String
    cpExtension = UCase(Right(cpPath, 7))
    
    Dim cpNode As IXMLDOMNode
    Set cpNode = ComponentsNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Dim cpAttr As IXMLDOMAttribute
    Dim cpSubNode As IXMLDOMNode
    
    Set cpAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "id", ""))
    cpAttr.nodeValue = swComponent.GetID()

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
    cpSubNode.Text = swComponent.ReferencedConfiguration
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "solving", ""))
    cpSubNode.Text = swComponent.Solving
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "suppression", ""))
    cpSubNode.Text = swComponent.GetSuppression2()
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "visible", ""))
    cpSubNode.Text = swComponent.Visible
    
    ExportComponentProps DOMDoc, cpNode, swComponent
    ExportMates swMath, DOMDoc, MatesNode, swComponent
    
    
Next


DOMDoc.loadXML Indent.Indent(DOMDoc.xml)
DOMDoc.Save swModel.GetPathName() + ".xml"
End Sub

Sub ExportComponentProps(DOMDoc As DOMDocument60, cpNode As IXMLDOMNode, swComponent As IComponent2)
Dim cpAttr As IXMLDOMAttribute
Dim cpSubNode As IXMLDOMNode
    
Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "transform", ""))

Dim j As Integer
For j = 0 To 15
    Dim cpTransformValueNode As IXMLDOMNode
    Set cpTransformValueNode = cpSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
    cpTransformValueNode.Text = swComponent.Transform2.ArrayData(j)
Next

Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "components", ""))

Dim swCpChildren As Variant
swCpChildren = swComponent.GetChildren()

Dim i As Integer
For i = LBound(swCpChildren) To UBound(swCpChildren)
    Dim swChild As IComponent2
    Set swChild = swCpChildren(i)
    
    Dim cpChildNode As IXMLDOMNode
    Set cpChildNode = cpSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Set cpAttr = cpChildNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "id", ""))
    cpAttr.nodeValue = swChild.GetID()
    
    ExportComponentProps DOMDoc, cpChildNode, swChild
Next

End Sub

Sub ExportMates(swMath As IMathUtility, DOMDoc As DOMDocument60, MatesNode As IXMLDOMNode, swComponent As IComponent2)
Dim mtAttr As IXMLDOMAttribute
Dim mtNode As IXMLDOMNode
Dim mtEntNode As IXMLDOMNode
Dim mtSubNode As IXMLDOMNode

Dim SingleMate As Variant
Dim swMates As Variant
swMates = swComponent.GetMates()

For Each SingleMate In swMates
    If TypeOf SingleMate Is SldWorks.Mate2 Then
            
        Dim swMate As IMate2
        Set swMate = SingleMate
        Dim swMateEnt As IMateEntity2
        Set swMateEnt = swMate.MateEntity(0)
        Dim swRefCp As IComponent2
        Set swRefCp = swMateEnt.ReferenceComponent
        
        Do Until swRefCp.GetParent() Is Nothing
            Set swRefCp = swRefCp.GetParent()
        Loop
        
        If swApp.IsSame(swRefCp, swComponent) <> swObjectSame Then GoTo MateSkip
    
        Set mtNode = DOMDoc.createNode(NODE_ELEMENT, "mate", "")
        
        Set mtAttr = mtNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "type", ""))
        mtAttr.nodeValue = swMate.Type
        
        Dim e As Integer
        For e = 0 To swMate.GetMateEntityCount() - 1
            Set swMateEnt = swMate.MateEntity(e)
            
            Set mtEntNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "entity", ""))
            
            Set mtAttr = mtEntNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "type", ""))
            mtAttr.nodeValue = swMateEnt.ReferenceType2
            
            Set swRefCp = swMateEnt.ReferenceComponent
            If swRefCp.GetID() = -1 Then GoTo MateSkip
            
            Dim mtRefCpID As String
            mtRefCpID = swRefCp.GetID()
            Dim swXForm As IMathTransform
            Set swXForm = swRefCp.Transform2
            Do Until swRefCp.GetParent() Is Nothing
                Set swRefCp = swRefCp.GetParent()
                mtRefCpID = swRefCp.GetID() & "/" & mtRefCpID
                If Not swRefCp.Transform2 Is Nothing Then Set swXForm = swXForm.Multiply(swRefCp.Transform2)
            Loop
            
            Set mtAttr = mtEntNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "component-id", ""))
            mtAttr.nodeValue = mtRefCpID
            
            Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "params", ""))
            
            Dim nPt(2) As Double
            Dim vPt As Variant
            Dim mtEntPt As IMathPoint
            Dim mtEntVec As IMathVector
            Dim mtParam(7) As Double
            Dim j As Integer
            
            For j = 0 To 2
                nPt(j) = swMateEnt.EntityParams(j)
            Next
            vPt = nPt
            Set mtEntPt = swMath.CreatePoint((vPt))
            Set mtEntPt = mtEntPt.MultiplyTransform(swXForm)
            For j = 0 To 2
                mtParam(j) = mtEntPt.ArrayData(j)
            Next
            For j = 0 To 2
                nPt(j) = swMateEnt.EntityParams(j + 3)
            Next
            vPt = nPt
            Set mtEntVec = swMath.CreateVector((vPt))
            Set mtEntVec = mtEntVec.MultiplyTransform(swXForm)
            For j = 0 To 2
                mtParam(j + 3) = mtEntVec.ArrayData(j)
            Next
            
            For j = 6 To 7
                mtParam(j) = swMateEnt.EntityParams(j)
            Next
            
            For j = 0 To 7
                Dim mtParamValueNode As IXMLDOMNode
                Set mtParamValueNode = mtSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
                mtParamValueNode.Text = mtParam(j)
            Next
            
        Next
                            
        ' Coincident
        Set mtSubNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "alignment", ""))
        mtSubNode.Text = swMate.Alignment
        
        MatesNode.appendChild mtNode
    End If
    
MateSkip:
Next
End Sub
