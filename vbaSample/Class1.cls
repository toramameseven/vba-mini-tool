VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private class_private As String
Public class_public As String
Dim class_dim As String


Public Function Public_Function() As String
    Public_Function = "Public_Function"
End Function

Function Non_Function() As String
    Non_Function = "Non_Function"
End Function

Private Function Private_Function() As String
    Private_Function = "Private_Function"
    Common
End Function
