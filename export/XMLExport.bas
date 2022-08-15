Attribute VB_Name = "XMLExport"
Dim swApp As Object

Dim swMath As IMathUtility

'アクティブなアセンブリをxmlにエクスポート
Sub main()
Set swApp = Application.SldWorks
Set swMath = swApp.GetMathUtility()

Dim swModel As ModelDoc2
Set swModel = swApp.ActiveDoc
If swModel.GetType() <> swDocASSEMBLY Then
    MsgBox "このファイルはアセンブリではありません"
    Exit Sub
End If

Dim swAsmDoc As IAssemblyDoc
Set swAsmDoc = swModel

Dim swConfMgr As IConfigurationManager
Set swConfMgr = swModel.ConfigurationManager
Dim swConf As IConfiguration
Set swConf = swConfMgr.ActiveConfiguration
Dim swRootCp As IComponent2
Set swRootCp = swConf.GetRootComponent3(True)

Dim DOMDoc As DOMDocument60
Set DOMDoc = New DOMDocument60
DOMDoc.appendChild DOMDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")

Dim cpAttr As IXMLDOMAttribute
Dim cpSubNode As IXMLDOMNode


Dim RootNode As IXMLDOMNode
Set RootNode = DOMDoc.appendChild(DOMDoc.createNode(NODE_ELEMENT, "assembly", ""))

Set cpAttr = RootNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "name", ""))
cpAttr.nodeValue = swRootCp.Name2 'GetSelectByIDString()=""

Set cpSubNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "path", ""))
cpSubNode.Text = swRootCp.GetPathName()

Dim ComponentsNode As IXMLDOMNode
Set ComponentsNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "components", ""))

Dim MatesNode As IXMLDOMNode
Set MatesNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "mates", ""))

Dim SingleComponent As Variant
For Each SingleComponent In swRootCp.GetChildren()
    Dim swComponent As IComponent2
    Set swComponent = SingleComponent

    Dim cpNode As IXMLDOMNode
    Set cpNode = ComponentsNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Set cpAttr = cpNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "name", ""))
    cpAttr.nodeValue = swComponent.GetSelectByIDString()

    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "path", ""))
    cpSubNode.Text = swComponent.GetPathName()
    
    Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "configuration", ""))
    cpSubNode.Text = swComponent.ReferencedConfiguration
    
    ExportComponentProps DOMDoc, cpNode, swComponent
    ExportMates swModel, DOMDoc, MatesNode, swComponent
    
    
Next


DOMDoc.loadXML Indent.Indent(DOMDoc.xml)
DOMDoc.Save swModel.GetPathName() + ".xml"
End Sub

'2階層以上のコンポーネントの出力
Sub ExportComponentProps(DOMDoc As DOMDocument60, cpNode As IXMLDOMNode, swComponent As IComponent2)
Dim cpAttr As IXMLDOMAttribute
Dim cpSubNode As IXMLDOMNode

Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "solving", ""))
cpSubNode.Text = swComponent.Solving

Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "suppression", ""))
cpSubNode.Text = swComponent.GetSuppression2()

Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "visible", ""))
cpSubNode.Text = swComponent.Visible

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
    cpAttr.nodeValue = swChild.GetSelectByIDString()
    
    ExportComponentProps DOMDoc, cpChildNode, swChild
Next

End Sub

'合致情報のエクスポート
Sub ExportMates(swModel As IModelDoc2, DOMDoc As DOMDocument60, MatesNode As IXMLDOMNode, swComponent As IComponent2)
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
        'Dim swMateEntRef As Object 'IMateReference
        
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
            'Set swMateEntRef = swMateEnt.Reference 'APIHelpの記述と違って選択したEntityなどが返る
            
            Set mtEntNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "entity", ""))
            
            Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "type", ""))
            mtSubNode.Text = swMateEnt.ReferenceType2
            
            Dim mtEntName As String
            Dim mtEntID(2) As String
            
            ' component.Name2
            Set mtAttr = mtEntNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "component", ""))
            'mtAttr.nodeValue = swRefCp.Name2
            mtAttr.nodeValue = swMateEnt.ReferenceComponent.GetSelectByIDString()
            
            ' Entityの名前
            'mtAttr.nodeValue = swMateEntRef.Name
            SelType.ExportEntityName swMateEnt.ReferenceType2, swMateEnt.Reference, mtEntName, mtEntID
            
            If mtEntName <> "" Then
                Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "name", ""))
                mtSubNode.Text = mtEntName
            End If
            
            Dim mtEntIDSingle As Variant
            For Each mtEntIDSingle In mtEntID
                If mtEntIDSingle <> "" Then
                    Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "id", ""))
                    mtSubNode.Text = mtEntIDSingle
                End If
            Next
            
            If mtEntName = "" Then
                Set mtSubNode = mtEntNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "params", ""))
                
                Dim nPt(2) As Double
                Dim vPt As Variant
                Dim mtEntPt As IMathPoint
                Dim ClosestPt As Variant
                Dim swClosestPt As IMathPoint
                Dim mtEntVec As IMathVector
                Dim mtParam(7) As Double
                Dim j As Integer
                
                For j = 0 To 7
                    mtParam(j) = swMateEnt.EntityParams(j)
                Next
                
                ' pointX, Y, Z
                For j = 0 To 2
                    nPt(j) = mtParam(j)
                Next
                vPt = nPt
                Set mtEntPt = swMath.CreatePoint((vPt))
                If Not swMateEnt.ReferenceComponent.Transform2 Is Nothing Then Set mtEntPt = mtEntPt.MultiplyTransform(swMateEnt.ReferenceComponent.Transform2.Inverse())

                If Not swMateEnt.Reference Is Nothing Then
                    If swMateEnt.ReferenceType2 = swSelEDGES Or swMateEnt.ReferenceType2 = swSelFACES Then
                        'EntityParamsの点が面上にないことがあり、その場合は面上で一番近い点にする
                        ClosestPt = swMateEnt.Reference.GetClosestPointOn(mtEntPt.ArrayData(0), mtEntPt.ArrayData(1), mtEntPt.ArrayData(2))
                        For j = 0 To 2
                            nPt(j) = ClosestPt(j)
                        Next
                        vPt = nPt
                        Set swClosestPt = swMath.CreatePoint((vPt))
                        If swClosestPt.Subtract(mtEntPt).GetLength() > mtParam(6) + 0.00001 Then Set mtEntPt = swClosestPt
                    ElseIf swMateEnt.ReferenceType2 = swSelVERTICES Then
                        mtEntPt = swMateEnt.Reference.GetPoint()
                    End If
                End If
                
                
                For j = 0 To 2
                    mtParam(j) = mtEntPt.ArrayData(j)
                Next
                
                ' vectorI, J ,K
                For j = 0 To 2
                    nPt(j) = mtParam(j + 3)
                Next
                vPt = nPt
                Set mtEntVec = swMath.CreateVector((vPt))
                If Not swMateEnt.ReferenceComponent.Transform2 Is Nothing Then Set mtEntVec = mtEntVec.MultiplyTransform(swMateEnt.ReferenceComponent.Transform2.Inverse())
                For j = 0 To 2
                    mtParam(j + 3) = mtEntVec.ArrayData(j)
                Next
                
                For j = 0 To 7
                    Dim mtParamValueNode As IXMLDOMNode
                    Set mtParamValueNode = mtSubNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
                    mtParamValueNode.Text = mtParam(j)
                Next
            End If
            
        Next
        
        ' Coincidentで使用
        Set mtSubNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "alignment", ""))
        mtSubNode.Text = swMate.Alignment
        
        MatesNode.appendChild mtNode
    End If
    
MateSkip:
Next
End Sub
