Attribute VB_Name = "OcctoXmlImporter"
Option Explicit

'param = {
    'typeOfPriPlan : "発販" or  "需調" or "連系線",
    'typeOfSecPlan : "発電"or"調達"or"販売"or"需要"or"連系線",
    'companyCode : "総計"orエリアコードを含めた事業者コード(typeOfPriPlanが発販と連系線の場合は空でOK)
    'geneBgCode:BGコード(typeOfPriPlanが需調の場合は不要),
    'gridCode:"系統コード"(typeOfPriPlanが連系線の場合は申込番号)
    'amountOrNot:"合計orブランク"(合計値を取りたいのか個別を取りたいのか)
'}

Function OcctoXmlImport(ByVal xmlPath As String, ByVal param As Object)
Dim paramKeys As Variant
Dim rightParamKeys As Variant
Dim chkLong As Variant
Dim returnArray(0 To 1) As Variant
Dim i As Long
Dim tempParamKeys As Variant
Dim tempParamValues As Variant
Dim returnValue(0 To 47) As Variant
tempParamKeys = Array("typeOfPriPlan", "typeOfSecPlan", "companyCode", "geneBgCode", "gridCode", "amountOrNot")
tempParamValues = Array("", "", "", "", "", "")
rightParamKeys = HashMakeFromArray(tempParamKeys, tempParamValues)

'─────paramのkeyが正しいかチェック─────
paramKeys = param.Keys
returnArray(0) = False
    For i = 0 To UBound(paramKeys)
        If rightParamKeys(1).exists(paramKeys(i)) = False Then
        returnArray(1) = "エラー：" & paramKeys(i) & "は不正なパラメータだよ。"
        OcctoXmlImport = returnArray
        Exit Function
        End If
    Next i
    If param.Item("gridCode") <> "" And param.Item("amountOrNot") <> "" Then
    returnArray(1) = "エラー：gridCodeとamountOrNotは同時には設定できないよ。"
    OcctoXmlImport = returnArray
    Exit Function
    End If

'─────チェック終わり─────

Dim XMLDocument As Object
Set XMLDocument = CreateObject("MSXML2.DOMDocument")
XMLDocument.async = False
Dim FileValue As Boolean
Dim SelNodeList As MSXML2.IXMLDOMNodeList
Dim Node  As IXMLDOMNode
Dim nodeString As String

Select Case param.Item("typeOfPriPlan")

Case "発販" '──────────────────────────────────────────────────────────
    Select Case param.Item("typeOfSecPlan")
        Case "発電"
            If param.Item("companyCode") = "総計" Then
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00012/JPMR00012/JPM00013/JPMR00013/JP06363"
            ElseIf param.Item("amountOrNot") = "合計" Then
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00014/JPMR00014[(JP06300='" & param.Item("geneBgCode") & "')]/JPM00015/JPMR00015/JP06307"
            Else
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00014/JPMR00014[(JP06300='" & param.Item("geneBgCode") & "')]/JPM00016/JPMR00016[(JP06186='" & param.Item("gridCode") & "')]/JPM00017/JPMR00017/JP06231"
            End If
            
        Case "販売"
            If param.Item("companyCode") = "総計" Then
            nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00018/JPMR00018/JPM00019/JPMR00019/JP06319"
            Else
            nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00018/JPMR00018/JPM00020/JPMR00020[(JP06366='" & param.Item("gridCode") & "')]/JPM00021/JPMR00021/JP06319"
            End If
        
        Case "調達"
            If param.Item("companyCode") = "総計" Then
            nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00022/JPMR00022/JPM00023/JPMR00023/JP06369"
            Else
            nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00022/JPMR00022/JPM00024/JPMR00024[(JP06366='" & param.Item("gridCode") & "')]/JPM00025/JPMR00025/JP06369"
            End If
        
        Case Else
            returnArray(1) = "エラー：思いもよらないtypeOfSecPlanでした。"
            OcctoXmlImport = returnArray
            Exit Function
    End Select


Case "需調"  '──────────────────────────────────────────────────────────
    Select Case param.Item("typeOfSecPlan")
        Case "需要"
            If param.Item("companyCode") = "総計" Then
            nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00010/JPMR00010/JPM00011/JPMR00011/JP06376"
            Else
            nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00022/JPMR00022[(JP06316='" & param.Item("companyCode") & "')]/JPM00023/JPMR00023/JPM00024/JPMR00024/JP06376"
            End If

        Case "販売"
            If param.Item("amountOrNot") = "合計" And param.Item("companyCode") = "総計" Then
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00018/JPMR00018/JPM00019/JPMR00019/JP06319"
            ElseIf param.Item("companyCode") = "総計" Then
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00018/JPMR00018/JPM00020/JPMR00020[(JP06366='" & param.Item("gridCode") & "')]/JPM00021/JPMR00021/JP06319"
            ElseIf param.Item("amountOrNot") = "合計" Then
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00022/JPMR00022[(JP06316='" & param.Item("companyCode") & "')]/JPM00031/JPMR00031/JPM00032/JPMR00032/JP06319"
            Else
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00022/JPMR00022[(JP06316='" & param.Item("companyCode") & "')]/JPM00031/JPMR00031/JPM00033/JPMR00033[(JP06366='" & param.Item("gridCode") & "')]/JPM00034/JPMR00034/JP06319"
            End If
        
        Case "調達"
            If param.Item("amountOrNot") = "合計" And param.Item("companyCode") = "総計" Then
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00014/JPMR00014/JPM00015/JPMR00015/JP06369"
            ElseIf param.Item("companyCode") = "総計" Then
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00014/JPMR00014/JPM00016/JPMR00016[(JP06366='" & param.Item("gridCode") & "')]/JPM00017/JPMR00017/JP06369"
            ElseIf param.Item("amountOrNot") = "合計" Then
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00022/JPMR00022[(JP06316='" & param.Item("companyCode") & "')]/JPM00027/JPMR00027/JPM00028/JPMR00028/JP06369"
            Else
                nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00022/JPMR00022[(JP06316='" & param.Item("companyCode") & "')]/JPM00027/JPMR00027/JPM00029/JPMR00029[(JP06366='" & param.Item("gridCode") & "')]/JPM00030/JPMR00030/JP06369"
            End If
        
        Case Else
            returnArray(1) = "エラー：思いもよらないtypeOfSecPlanでした。"
            OcctoXmlImport = returnArray
            Exit Function
    End Select

Case "連系線" '──────────────────────────────────────────────────────────
    Select Case param.Item("typeOfSecPlan")
        Case "連系線"
            nodeString = "SBD-MSG/JPMGRP/JPTRM/JPM00010/JPMR00010[(JP06185='" & param.Item("gridCode") & "')]/JPM00013/JPMR00013/JP06228"
        Case Else
            returnArray(1) = "エラー：思いもよらないtypeOfSecPlanでした。"
            OcctoXmlImport = returnArray
            Exit Function
    End Select

Case Else '──────────────────────────────────────────────────────────
    returnArray(1) = "エラー：思いもよらないtypeOfPriPlanでした。"
    OcctoXmlImport = returnArray
    Exit Function
End Select

'─────xml読み込み─────
FileValue = XMLDocument.Load(xmlPath)
    If FileValue = False Then
        returnArray(1) = "エラー：xmlPathが存在しないよ。"
        OcctoXmlImport = returnArray
        Exit Function
    End If

Set SelNodeList = XMLDocument.SelectNodes(nodeString)
    If SelNodeList.Length = 0 Then
    returnArray(1) = "エラー：指定されたbgCode、gridCode、companyCodeのどれかが存在しなかったよ。"
    OcctoXmlImport = returnArray
    Exit Function
    End If

i = 0
    For Each Node In SelNodeList
    returnValue(i) = Val(Node.ChildNodes(0).Text)
    i = i + 1
Next

Set XMLDocument = Nothing
Set SelNodeList = Nothing
Set Node = Nothing

returnArray(0) = True
returnArray(1) = returnValue
OcctoXmlImport = returnArray

End Function


Function HashMakeFromArray(ByVal keysArray As Variant, ByVal valuesArray As Variant)
Dim returnHash As Object
Dim returnValue(0 To 1) As Variant
Dim i As Long
returnValue(0) = False
Set returnHash = CreateObject("Scripting.Dictionary")
'─────2つのArrayの個数が違わないかチェック─────
If UBound(keysArray) <> UBound(valuesArray) Then
returnValue(1) = "くれた配列の数が一致しないよ"
HashMakeFromArray = returnValue
Exit Function
End If
'─────チェック終わり─────
For i = 0 To UBound(keysArray)
returnHash.Add keysArray(i), valuesArray(i)
Next i
returnValue(0) = True
Set returnValue(1) = returnHash
HashMakeFromArray = returnValue
End Function


