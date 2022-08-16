Attribute VB_Name = "XMLImport"
Dim swApp As Object

Dim swMath As IMathUtility

Dim cpNameReplaceList As Object

Function cpNameReplace(cpName As String) As String
Dim oldName As Variant
For Each oldName In cpNameReplaceList
    cpName = Replace(cpName, oldName, cpNameReplaceList(oldName))
Next
cpNameReplace = cpName
End Function

'xmlからアセンブリ生成
Sub main()
Set swApp = Application.SldWorks
Set swMath = swApp.GetMathUtility()

Set cpNameReplaceList = CreateObject("Scripting.Dictionary")

'名前の変更を有効にする
Dim ExtRefUpdateCompNamesDefault As Boolean
ExtRefUpdateCompNamesDefault = swApp.GetUserPreferenceToggle(swExtRefUpdateCompNames)
swApp.SetUserPreferenceToggle swExtRefUpdateCompNames, False


Dim OpenFile As Variant
OpenFile = swApp.ActiveDoc.GetPathName() + ".xml"

'新規アセンブリ
Dim swTemplate As String
swTemplate = swApp.GetDocumentTemplate(swDocASSEMBLY, "", 0, 0, 0)
Dim swModel As ModelDoc2
Set swModel = swApp.NewDocument(swTemplate, 0, 0, 0)

Dim swAsmDoc As IAssemblyDoc
Set swAsmDoc = swModel

Dim swSelMgr As ISelectionMgr
Set swSelMgr = swModel.SelectionManager

Dim swConfMgr As IConfigurationManager
Set swConfMgr = swModel.ConfigurationManager
Dim swConf As IConfiguration
Set swConf = swConfMgr.ActiveConfiguration
Dim swRootCp As IComponent2
Set swRootCp = swConf.GetRootComponent3(True)

Dim DOMDoc As DOMDocument60
Set DOMDoc = New DOMDocument60
DOMDoc.Load OpenFile

Dim RootCpName As String
RootCpName = DOMDoc.selectSingleNode("/assembly/@name").Text

Dim cpNode As IXMLDOMElement
For Each cpNode In DOMDoc.selectNodes("/assembly/components/component")
    Dim cpName As String
    Dim cpPath As String
    Dim cpConfiguration As String
   
    cpName = cpNode.getAttribute("name")
    cpPath = cpNode.selectSingleNode("path").Text
    cpConfiguration = cpNode.selectSingleNode("configuration").Text
    
    Dim swComponent As IComponent2
    Set swComponent = swAsmDoc.AddComponent5(cpPath, swAddComponentConfigOptions_CurrentSelectedConfig, "", True, cpConfiguration, 0, 0, 0)
    cpNameReplaceList.Add cpName, swComponent.GetSelectByIDString()
    swComponent.Select4 False, Nothing, False
    'swComponent.Name2 = cpName
    Debug.Print swComponent.GetSelectByIDString()
   
    'ApplyComponentProps cpNode, swComponent, swAsmDoc
    'HideAllComponent swComponent
    
Next


Dim mtNode As IXMLDOMElement
For Each mtNode In DOMDoc.selectNodes("/assembly/mates/mate")
    swModel.ClearSelection2 True

    Dim mtEntityNodes As Object
    Set mtEntityNodes = mtNode.selectNodes("entity")
    
    Dim mtEntNode As IXMLDOMElement
    Dim i As Integer
    'For Each mtEntNode In mtNode.selectNodes("entity")
    For i = 0 To mtEntityNodes.Length - 1
        Set mtEntNode = mtEntityNodes(i)
        SelType.SelectEntity mtEntNode, swAsmDoc, swMath
    Next
    Debug.Print swSelMgr.GetSelectedObjectCount2(1)
    
            
    
    'For i = 0 To mtEntityNodes.Length - 1
    '    TargetEntities(i).Select4 True, swSelectData
    'Next
    
    Dim mtType As Integer
    mtType = mtNode.selectSingleNode("type").Text
    
    Dim mtData As IMateFeatureData
    Set mtData = swAsmDoc.CreateMateData(mtType)
    
    If mtType = swMateCOINCIDENT Then
        Dim mtDataCasted As ICoincidentMateFeatureData
        Set mtDataCasted = mtData
    
        mtDataCasted.MateAlignment = mtNode.selectSingleNode("alignment").Text
    End If
    swAsmDoc.CreateMate mtDataCasted
    
    'For i = 0 To mtEntityNodes.Length - 1
    '    Set mtEntNode = mtEntityNodes(i)
    '    SelType.HideEntityComponent mtEntNode, swAsmDoc
    'Next

Next

Dim swChild As Variant
Dim swChildren As Variant
swChildren = swRootCp.GetChildren()
For Each swChild In swChildren
    Dim swChildCp As IComponent2
    Set swChildCp = swChild
    Dim cpElement As IXMLDOMElement
    For Each cpElement In DOMDoc.selectNodes("/assembly/components/component")
        Dim cpElementName As String
        cpElementName = cpNameReplace(cpElement.getAttribute("name"))
        If cpElementName = swChildCp.GetSelectByIDString() Then
            Debug.Print swComponent.GetSelectByIDString()
            ApplyComponentProps cpElement, swChildCp, swAsmDoc
            Exit For
        End If
    Next
Next

'設定を戻す
swApp.SetUserPreferenceToggle swExtRefUpdateCompNames, ExtRefUpdateCompNamesDefault
End Sub

Sub ApplyComponentProps(cpNode As IXMLDOMElement, swComponent As IComponent2, swAsmDoc As IAssemblyDoc)
Dim cpSolving As Integer
Dim cpVisible As Boolean
Dim cpSuppression As Integer
Dim cpTransformNodes As IXMLDOMNodeList
Dim cpChildren As IXMLDOMNodeList

cpSolving = cpNode.selectSingleNode("solving").Text
cpVisible = cpNode.selectSingleNode("visible").Text
cpSuppression = cpNode.selectSingleNode("suppression").Text

swComponent.Select4 False, Nothing, False
swAsmDoc.CompConfigProperties6 cpSuppression, cpSolving, cpVisible, False, "", False, False, False

Set cpTransformNodes = cpNode.selectNodes("transform/value")
Set cpChildren = cpNode.selectNodes("components/component")

Dim TransformArray(15) As Double
Dim j As Integer
For j = 0 To 12
    TransformArray(j) = cpTransformNodes(j).Text
Next
swComponent.Transform2 = swMath.CreateTransform(TransformArray)

Dim swChild As Variant
Dim swChildren As Variant
swChildren = swComponent.GetChildren()
For Each swChild In swChildren
    Dim swChildCp As IComponent2
    Set swChildCp = swChild
    Dim cpElement As IXMLDOMElement
    For Each cpElement In cpChildren
        Dim cpElementName As String
        cpElementName = cpNameReplace(cpElement.getAttribute("name"))
        If cpElementName = swChildCp.GetSelectByIDString() Then
            Debug.Print swComponent.GetSelectByIDString()
            ApplyComponentProps cpElement, swChildCp, swAsmDoc
            Exit For
        End If
    Next
Next
    
End Sub

Sub HideAllComponent(swComponent As IComponent2)
Dim cpChildren As IXMLDOMNodeList

swComponent.SetVisibility swComponentHidden, swThisConfiguration, ""
If swComponent.GetSuppression2() = swComponentSuppressed Then swComponent.SetSuppression2 swComponentResolved

Dim swChild As Variant
Dim swChildren As Variant
swChildren = swComponent.GetChildren()
For Each swChild In swChildren
    Dim swChildCp As IComponent2
    Set swChildCp = swChild
    HideAllComponent swChildCp
Next
    
End Sub

