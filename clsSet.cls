VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class Module: clsSet
' Uses VBA Dictionary class to mimic the behavior of a set
' If getting errors, try referencing the Microsoft Scripting Runtime library under Tools
'
' Example:
'
' Private numSet As clsSet
' Set numSet = New clsSet
'

Private dict As Object
Private key As Variant

Private Sub Class_Initialize()

    Set dict = CreateObject("Scripting.Dictionary")
    
End Sub

Public Sub Add(variable As Variant)

    If Not dict.Exists(variable) Then
    
        dict.Add variable, 0
        
    End If
    
End Sub

Public Function Exists(variable As Variant) As Boolean

    Exists = dict.Exists(variable)
    
End Function

Public Sub Remove(variable As Variant)

    dict.Remove variable
    
End Sub

Private Function getVar(variable As Variant)

    getVar = dict(variable)

End Function

Property Get items() As Variant

    items = dict.keys

End Property

Private Sub Class_Terminate()

    Set dict = Nothing
    
End Sub

