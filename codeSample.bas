Attribute VB_Name = "codeSample"
Option Explicit

'param = {
    'typeOfPriPlan : "発販" or  "需調" or "連系線",
    'typeOfSecPlan : "発電"or"調達"or"販売"or"需要",
    'companyCode : "合計"orエリアコードを含めた事業者コード(typeOfPriPlanが発販と連系線の場合は不要)
    'bgCode:BGコード(typeOfPriPlanが需調の場合は不要),
    'gridCode:"系統コード"(typeOfPriPlanが連系線の場合は申込番号)
    'amountOrNot:"合計orブランク"(合計値を取りたいのか個別を取りたいのか)
'}

Sub test()
Dim result As Variant
Dim param As Variant
Dim paramKeys As Variant
Dim paramValues As Variant
Dim xmlPath As String
Dim i As Long

paramKeys = Array("typeOfPriPlan", "typeOfSecPlan", "companyCode", "bgCode", "gridCode", "amountOrNot")
paramValues = Array("発販", "発電", "", "LZ999", "", "合計")
param = HashMakeFromArray(paramKeys, paramValues)
xmlPath = "C:\plans\W6_0150_20171211_00_99992_2.xml"

result = OcctoXmlImport(xmlPath, param(1))
Debug.Print result(0)
ThisWorkbook.Sheets("Sheet1").Range("H1:H48").Value = WorksheetFunction.Transpose(result(1))


End Sub
