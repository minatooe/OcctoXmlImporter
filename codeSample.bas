Attribute VB_Name = "codeSample"
Option Explicit

'param = {
    'typeOfPriPlan : "発販" or  "需調" or "連系線",　#string#
    'typeOfSecPlan : "発電"or"調達"or"販売"or"需要"or"連系線",　#string#
    'companyCode : "総計"orエリアコードを含めた事業者コード(typeOfPriPlanが発販と連系線の場合は不要),　#string#
    'geneBgCode:"発電BGコード"(typeOfPriPlanが需調か連系線の場合は不要),　#string#
    'gridCode:"系統コードor取引先BGコードor申込番号(typeOfPriPlanが連系線の場合)",　#string#
    'amountOrNot:"合計orブランク"(typeOfSecPlanで指定した計画の合計値を取りたい場合は"合計"を指定してください。)　#string#
'}

Sub test()
Dim result As Variant
Dim param As Variant
Dim paramKeys As Variant
Dim paramValues As Variant
Dim xmlPath As String
Dim i As Long

Dim typeOfPriPlan, typeOfSecPlan, companyCode, geneBgCode, gridCode, amountOrNot As String
typeOfPriPlan = "発販"
typeOfSecPlan = "調達"
companyCode = ""
geneBgCode = ""
gridCode = "ZZ999"
amountOrNot = ""

paramKeys = Array("typeOfPriPlan", "typeOfSecPlan", "companyCode", "geneBgCode", "gridCode", "amountOrNot") 'paramKeysは変えないことを推奨
paramValues = Array(typeOfPriPlan, typeOfSecPlan, companyCode, geneBgCode, gridCode, amountOrNot)
param = HashMakeFromArray(paramKeys, paramValues)
xmlPath = "c:\plan\W6_0150_20171201_00_99999_9.xml"

result = OcctoXmlImport(xmlPath, param(1))
If result(0) Then
ThisWorkbook.Sheets("Sheet1").Range("H1:H48").Value = WorksheetFunction.Transpose(result(1))
Else
MsgBox (result(1))
End If
End Sub
