Attribute VB_Name = "codeSample"
Option Explicit

'param = {
    'typeOfPriPlan : "����" or  "����" or "�A�n��",�@#string#
    'typeOfSecPlan : "���d"or"���B"or"�̔�"or"���v"or"�A�n��",�@#string#
    'companyCode : "���v"or�G���A�R�[�h���܂߂����Ǝ҃R�[�h(typeOfPriPlan�����̂ƘA�n���̏ꍇ�͕s�v),�@#string#
    'geneBgCode:"���dBG�R�[�h"(typeOfPriPlan���������A�n���̏ꍇ�͕s�v),�@#string#
    'gridCode:"�n���R�[�hor�����BG�R�[�hor�\���ԍ�(typeOfPriPlan���A�n���̏ꍇ)",�@#string#
    'amountOrNot:"���vor�u�����N"(typeOfSecPlan�Ŏw�肵���v��̍��v�l����肽���ꍇ��"���v"���w�肵�Ă��������B)�@#string#
'}

Sub test()
Dim result As Variant
Dim param As Variant
Dim paramKeys As Variant
Dim paramValues As Variant
Dim xmlPath As String
Dim i As Long

Dim typeOfPriPlan, typeOfSecPlan, companyCode, geneBgCode, gridCode, amountOrNot As String
typeOfPriPlan = "����"
typeOfSecPlan = "���B"
companyCode = ""
geneBgCode = ""
gridCode = "ZZ999"
amountOrNot = ""

paramKeys = Array("typeOfPriPlan", "typeOfSecPlan", "companyCode", "geneBgCode", "gridCode", "amountOrNot") 'paramKeys�͕ς��Ȃ����Ƃ𐄏�
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
