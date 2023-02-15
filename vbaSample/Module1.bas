Attribute VB_Name = "Module1"
Option Explicit

Dim Module1_Dim_Val
Public Module1_Public_Val
Private Module1_Private_Val


Public Sub Module1()
    'Module2_Dim_Val = 10 '' can not be accessed form other modules.
    Module2_Public_Val = 100
End Sub

Public Sub Common()
    'Module1_Dim_Val = 10
End Sub

Public Sub Log(ByRef msg As String)
    Debug.Print msg
End Sub
