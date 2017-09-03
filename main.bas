Attribute VB_Name = "Module1"
Option Explicit

Sub Main()

    If App.PrevInstance Then
        Exit Sub
    End If
    
    Form1.Show
    
End Sub

