Attribute VB_Name = "XMLExport"
Dim swApp As Object

'アクティブなアセンブリをxmlにエクスポート
Sub main()
Set swApp = Application.SldWorks

Dim swMath As IMathUtility
Set swMath = swApp.GetMathUtility()

Dim swModel As ModelDoc2
Set swModel = swApp.ActiveDoc
If swModel.GetType() <> swDocASSEMBLY Then
    MsgBox "このファイルはアセンブリではありません"
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

Dim SingleComponent As Variant
For Each SingleComponent In swAsmDoc.GetComponents(True)
    Dim swComponent As IComponent2
    Set swComponent = SingleComponent

    Dim cpPath As String
    cpPath = swComponent.GetPathName()
    Dim cpExtension As String
    cpExtension = UCase(Right(cpPath, 7))
    
    Dim cpNode As IXMLDOMNode
    Set cpNode = ComponentsNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Dim cpAttr As IXMLDOMAttribute
    Dim cpSubNode As IXMLDOMNode
    
    Set cpAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "name", ""))
    cpAttr.nodeValue = swComponent.Name2

    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "path", ""))
    cpSubNode.Text = cpPath

    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "type", ""))
    'cpSubNode.Text = cpModel.GetType()
    If cpExtension = ".SLDASM" Then
        cpSubNode.Text = swDocASSEMBLY
    ElseIf cpExtension = ".SLDPRT" Then
        cpSubNode.Text = swDocPART
    Else
        MsgBox "拡張子「" & cpExtension & "」は非対応です"
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
    ExportMates swMath, swModel, DOMDoc, MatesNode, swComponent
    
    
Next


DOMDoc.loadXML Indent.Indent(DOMDoc.xml)
DOMDoc.Save swModel.GetPathName() + ".xml"
End Sub

'2階層以上のコンポーネントに必要な情報は名前(Name2)と位置(Transform2)のみ
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


Dim Child As Variant
For Each Child In swComponent.GetChildren()
    Dim swChild As IComponent2
    Set swChild = Child
    
    Dim cpChildNode As IXMLDOMNode
    Set cpChildNode = cpSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Set cpAttr = cpChildNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "name", ""))
    cpAttr.nodeValue = swChild.Name2
    
    ExportComponentProps DOMDoc, cpChildNode, swChild
Next

End Sub

'合致情報のエクスポート
Sub ExportMates(swMath As IMathUtility, swModel As IModelDoc2, DOMDoc As DOMDocument60, MatesNode As IXMLDOMNode, swComponent As IComponent2)
Dim mtAttr As IXMLDOMAttribute
Dim mtNode As IXMLDOMNode
Dim mtEntNode As IXMLDOMNode
Dim mtSubNode As IXMLDOMNode

Dim swSelMgr As ISelectionMgr
Set swSelMgr = swModel.SelectionManager

Dim SingleMate As Variant
Dim swMates As Variant
swMates = swComponent.GetMates()

If IsEmpty(swMates) Then Exit Sub

For Each SingleMate In swMates
    If TypeOf SingleMate Is SldWorks.Mate2 Then
            
        Dim swMate As IMate2
        Set swMate = SingleMate
        
        Dim swMateEnt As IMateEntity2
        Dim swMateEntRef As Object 'IMateReference
        Dim swRefCp As IComponent2
        
        '2個のComponentから同じMateにアクセスでき、
        '片方からみるとMateEntity(0)、もう片方からはMateEntity(1)が自身に属することになるので、
        'MateEntity(0)が自分に属する場合のみエクスポートすることにする
        Set swMateEnt = swMate.MateEntity(0)
        Set swRefCp = swMateEnt.ReferenceComponent
        Do Until swRefCp.GetParent() Is Nothing
            Set swRefCp = swRefCp.GetParent()
        Loop
        If swApp.IsSame(swRefCp, swComponent) <> swObjectSame Then GoTo MateSkip
    
        Set mtNode = DOMDoc.createNode(NODE_ELEMENT, "mate", "")
        
        Set mtSubNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "type", ""))
        mtSubNode.Text = swMate.Type
        
        Dim e As Integer
        For e = 0 To swMate.GetMateEntityCount() - 1
            Set swMateEnt = swMate.MateEntity(e)
            Set swMateEntRef = swMateEnt.Reference
            
            Set mtEntNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "entity", ""))
            
            Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "type", ""))
            mtSubNode.Text = swMateEnt.ReferenceType2
            
            ' Entityの名前
            Set mtAttr = mtEntNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "name", ""))
            mtAttr.nodeValue = swMateEntRef.Name
            
            
            Set swRefCp = swMateEnt.ReferenceComponent
            
            ' component.Name2
            Set mtAttr = mtEntNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "component", ""))
            mtAttr.nodeValue = swRefCp.Name2
            
            ' Rootアセンブリから見た該当ComponentのTransform
            Dim swXForm As IMathTransform
            Set swXForm = swRefCp.Transform2
            Do Until swRefCp.GetParent() Is Nothing
                Set swRefCp = swRefCp.GetParent()
                If Not swRefCp.Transform2 Is Nothing Then Set swXForm = swXForm.Multiply(swRefCp.Transform2)
            Loop
            
            
            Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "params", ""))
            
            Dim nPt(2) As Double
            Dim vPt As Variant
            Dim mtEntPt As IMathPoint
            Dim mtEntVec As IMathVector
            Dim mtParam(7) As Double
            Dim j As Integer
            
            ' pointX, Y, Z
            For j = 0 To 2
                nPt(j) = swMateEnt.EntityParams(j)
            Next
            vPt = nPt
            Set mtEntPt = swMath.CreatePoint((vPt))
            Set mtEntPt = mtEntPt.MultiplyTransform(swXForm)
            For j = 0 To 2
                mtParam(j) = mtEntPt.ArrayData(j)
            Next
            
            ' vectorI, J ,K
            For j = 0 To 2
                nPt(j) = swMateEnt.EntityParams(j + 3)
            Next
            vPt = nPt
            Set mtEntVec = swMath.CreateVector((vPt))
            Set mtEntVec = mtEntVec.MultiplyTransform(swXForm)
            For j = 0 To 2
                mtParam(j + 3) = mtEntVec.ArrayData(j)
            Next
            
            ' radius1, 2
            For j = 6 To 7
                mtParam(j) = swMateEnt.EntityParams(j)
            Next
            
            For j = 0 To 7
                Dim mtParamValueNode As IXMLDOMNode
                Set mtParamValueNode = mtSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
                mtParamValueNode.Text = mtParam(j)
            Next
            
        Next
        
        ' Coincidentで使用
        Set mtSubNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "alignment", ""))
        mtSubNode.Text = swMate.Alignment
        
        MatesNode.appendChild mtNode
    End If
    
MateSkip:
Next
End Sub
