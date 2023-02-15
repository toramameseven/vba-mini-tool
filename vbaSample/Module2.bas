Attribute VB_Name = "Module2"
Option Explicit

Dim Module2_Dim_Val
Public Module2_Public_Val
Private Module2_Private_Val


Public Sub Module2()
    'Module1_Dim_Val = 10
End Sub

Public Sub Common()
    'Module1_Dim_Val = 10
End Sub

Public Sub Log(ByRef msg As String)
    Dim ccc As Class1
    Set ccc = New Class1
    ccc.class_public = "aaaa"
    ccc.Non_Function
    ccc.Public_Function
    Debug.Print msg
End Sub

