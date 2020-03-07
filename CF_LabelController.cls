VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CF_LabelController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Label As MSForms.Label
Attribute Label.VB_VarHelpID = -1

Private CFP As CF_Calclation

Property Get setControlClass(ByRef con As MSForms.Control) As CF_LabelController
    
    '/*
    '* Class設定
    
    Set Label = con
    
    Set setControlClass = Me
    
    Set CFP = New CF_Calclation
    
End Property

Private Sub Label_Click()
    
    '/*
    '* クリック時動作
    
    Call LabelClickEvent(Label)
    
End Sub

Private Sub Label_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    '/*
    '* クリック時動作
    
    Call LabelClickEvent(Label)
    
End Sub

Private Sub LabelClickEvent(con As MSForms.Control)
    
    '/*
    '* クリック時動作するプログラム
    
    CFP.Calclate con.Caption
    
End Sub
