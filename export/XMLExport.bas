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

Dim DOMDoc As DOMDocument60
Set DOMDoc = New DOMDocument60
DOMDoc.appendChild DOMDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")

Dim cpAttr As IXMLDOMAttribute
Dim cpSubNode As IXMLDOMNode


Dim RootNode As IXMLDOMNode
Set RootNode = DOMDoc.appendChild(DOMDoc.createNode(NODE_ELEMENT, "assembly", ""))

Dim ConfsNode As IXMLDOMNode
Set ConfsNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "configurations", ""))

Dim swConf As IConfiguration
Dim ConfNames As Variant
ConfNames = swModel.GetConfigurationNames()
Dim ConfName As Variant
For Each ConfName In ConfNames
    Set swConf = swModel.GetConfigurationByName(ConfName)

    Dim cfNode As IXMLDOMNode
    Set cfNode = ConfsNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "configuration", ""))
    
    Set cpAttr = cfNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "name", ""))
    cpAttr.nodeValue = swConf.Name
    
    Set cpAttr = cfNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "parent", ""))
    If swConf.IsDerived Then cpAttr.nodeValue = swConf.GetParent().Name
Next

Dim swConfMgr As IConfigurationManager
Set swConfMgr = swModel.ConfigurationManager
Set swConf = swConfMgr.ActiveConfiguration
Dim swRootCp As IComponent2
Set swRootCp = swConf.GetRootComponent3(True)

Set cpAttr = RootNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "name", ""))
cpAttr.nodeValue = swRootCp.Name2 'GetSelectByIDString()=""

Set cpSubNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "path", ""))
cpSubNode.Text = swRootCp.GetPathName()

Dim ComponentsNode As IXMLDOMNode
Set ComponentsNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "toplevel", ""))

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

Next


'子コンポーネントすべてのTransformの出力
Set ComponentsNode = RootNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "transform", ""))
ExportComponentProps DOMDoc, ComponentsNode, swRootCp


For Each ConfName In ConfNames
    Dim swConfName As String
    swConfName = ConfName

    Set swConf = swModel.GetConfigurationByName(swConfName)
    swModel.ShowConfiguration2 swConfName
    
    '各コンフィギュレーションで抑制、固定、フレキシブル、参照コンフィギュレーションの情報を出力
    Set swRootCp = swConf.GetRootComponent3(True)
    
    For Each SingleComponent In swRootCp.GetChildren()
        Set swComponent = SingleComponent

        Set cpNode = RootNode.selectSingleNode("toplevel/component[@name=""" & swComponent.GetSelectByIDString() & """]")

        Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "reference", ""))
        cpSubNode.Text = swComponent.ReferencedConfiguration
        Set cpAttr = cpSubNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "configuration", ""))
        cpAttr.nodeValue = swConfName
        
        Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "suppression", ""))
        cpSubNode.Text = swComponent.GetSuppression2()
        Set cpAttr = cpSubNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "configuration", ""))
        cpAttr.nodeValue = swConfName
        
        Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "solving", ""))
        cpSubNode.Text = swComponent.Solving
        Set cpAttr = cpSubNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "configuration", ""))
        cpAttr.nodeValue = swConfName
        
        If swComponent.IsFixed Then
            Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "fixed", ""))
            Set cpAttr = cpSubNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "configuration", ""))
            cpAttr.nodeValue = swConfName
        End If
        
        ExportMates swModel, DOMDoc, MatesNode, swComponent, swConfName
    Next

Next


DOMDoc.loadXML Indent.Indent(DOMDoc.xml)
DOMDoc.Save swModel.GetPathName() + ".xml"
End Sub

'2階層以上のコンポーネントの出力
Sub ExportComponentProps(DOMDoc As DOMDocument60, cpNode As IXMLDOMNode, swComponent As IComponent2)
Dim cpAttr As IXMLDOMAttribute
Dim cpSubNode As IXMLDOMNode
    
'Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "solving", ""))
'cpSubNode.Text = swComponent.Solving

'Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "suppression", ""))
'cpSubNode.Text = swComponent.GetSuppression2()

'Set cpSubNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "visible", ""))
'cpSubNode.Text = swComponent.Visible

Dim Child As Variant
For Each Child In swComponent.GetChildren()
    Dim swChild As IComponent2
    Set swChild = Child
    
    Dim cpChildNode As IXMLDOMNode
    Set cpChildNode = cpNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "component", ""))
    
    Set cpAttr = cpChildNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "name", ""))
    cpAttr.nodeValue = swChild.GetSelectByIDString()

    Dim j As Integer
    For j = 0 To 15
        Dim cpTransformValueNode As IXMLDOMNode
        Set cpTransformValueNode = cpChildNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "value", ""))
        cpTransformValueNode.Text = swChild.Transform2.ArrayData(j)
    Next
    
    ExportComponentProps DOMDoc, cpNode, swChild
Next

End Sub

'合致情報のエクスポート
Sub ExportMates(swModel As IModelDoc2, DOMDoc As DOMDocument60, MatesNode As IXMLDOMNode, swComponent As IComponent2, ConfName As String)
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
    
        Set mtNode = DOMDoc.createNode(NODE_ELEMENT, "mate", "")
        
        Set mtAttr = mtNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "name", ""))
        mtAttr.nodeValue = swMate.Name
                
        Set mtSubNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "type", ""))
        mtSubNode.Text = swMate.Type
        

        Dim e As Integer
        For e = 0 To swMate.GetMateEntityCount() - 1
            Set swMateEnt = swMate.MateEntity(e)
            'Set swMateEntRef = swMateEnt.Reference 'APIHelpの記述と違って選択したEntityなどが返る
            
            If swMateEnt.Reference Is Nothing Then GoTo MateSkip
            
            Set mtEntNode = DOMDoc.createNode(NODE_ELEMENT, "entity", "")
            
            Dim ExportState As Boolean
            ExportState = SelType.ExportEntity(swMateEnt, mtEntNode, DOMDoc, swMath)
            If ExportState Then mtNode.appendChild mtEntNode
            
        Next
        
        ' Coincidentで使用
        Set mtSubNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "alignment", ""))
        mtSubNode.Text = swMate.Alignment
        
                
        If Not MatesNode.selectSingleNode("mate[@name=""" & swMate.Name & """]") Is Nothing Then
            Set mtNode = MatesNode.selectSingleNode("mate[@name=""" & swMate.Name & """]")
            If mtNode.selectSingleNode("active[@configuration=""" & ConfName & """]") Is Nothing Then
                Set mtSubNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "active", ""))
                Set mtAttr = mtSubNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "configuration", ""))
                mtAttr.nodeValue = ConfName
            End If
            GoTo MateSkip
        End If

        Set mtSubNode = mtNode.appendChild(DOMDoc.createNode(NODE_ELEMENT, "active", ""))
        Set mtAttr = mtSubNode.Attributes.setNamedItem(DOMDoc.createNode(NODE_ATTRIBUTE, "configuration", ""))
        mtAttr.nodeValue = ConfName
        
        MatesNode.appendChild mtNode
        
    End If
    
MateSkip:
Next
End Sub
