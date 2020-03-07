VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalclaterForm 
   Caption         =   "CalcForm"
   ClientHeight    =   2400
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2964
   OleObjectBlob   =   "CalclaterForm.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "CalclaterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ControlCollection As New Collection

Private Sub UserForm_initialize()

    '/*
    '* UserForm起動時実行

    Call formInitialize
    Call setControls
    
End Sub

Private Sub formInitialize()
    
    Dim i As Long, j As Long
    Dim labelTop As Long, labelLeft As Long
    
    '/*
    '* FormSize Setting
    
    With Me
        
        .Height = 390: .Width = (.Height / 3) * 3.3
        
        .BackColor = &H8000000F
        
    End With
    
    '/*
    '* 1~9数字キー配置
    
    For j = 1 To 3
    
        For i = 1 To 3
            
            labelTop = 130 + ((j - 1) * 55)
            labelLeft = 70 + ((i - 1) * 75)
            
            Call ControlAdd("Label", (9 - (j * 3) + i), _
                             i, j, _
                             labelTop, labelLeft, _
                             40, 60, _
                             2, &HC5D8EB)
            
        Next i
    
    Next j
    
    '/*
    '* 0 と　「.」のキー
    
    labelTop = 130 + ((j - 1) * 55)
    labelLeft = 70
            
            Call ControlAdd("Label", "0", _
                             i, j, _
                             labelTop, labelLeft, _
                             40, (60 * 2 + 15), _
                             2, &HC5D8EB)
    
    labelTop = 130 + ((j - 1) * 55)
    labelLeft = 70 + (2 * 75)
            
            Call ControlAdd("Label", ".", _
                             i, j, _
                             labelTop, labelLeft, _
                             40, 60, _
                             2, &HC5D8EB)
                             
    '/*
    '* / * + - C D = のキー
    
    Dim n As Long: n = 1
    
    For i = 1 To 4
            
            labelTop = 130 - 55
            labelLeft = 70 + ((i - 1) * 75)
            
            Call ControlAdd("Label", captionList(n), _
                             i, j, _
                             labelTop, labelLeft, _
                             40, 60, _
                             2, &HD1FCED)
            
            n = n + 1
            
    Next i
                             
    For j = 1 To 2
            
            labelTop = 130 + ((j - 1) * 55)
            labelLeft = 70 + (3 * 75)
            
            Call ControlAdd("Label", captionList(n), _
                             i, j, _
                             labelTop, labelLeft, _
                             40, 60, _
                             2, &HD1FCED)
                             
            n = n + 1
            
    Next j
                             
            labelTop = 130 + (2 * 55)
            labelLeft = 70 + (3 * 75)
            
            Call ControlAdd("Label", captionList(n), _
                             i, j, _
                             labelTop, labelLeft, _
                             (40 * 2 + 15), 60, _
                             2, &HD1FCED)
                             
            n = n + 1
                             
    '/*
    '* 計算結果
    
            labelTop = 130 - (2 * 55)
            labelLeft = 70
            
            Call ControlAdd("CalcStatus", "", _
                             0, 0, _
                             labelTop, labelLeft, _
                             40, (60 * 4 + 45), _
                             3, &HFFFFFF)
    
    
End Sub

Private Sub ControlAdd(conName As String, conCaption As String, _
                       conColumn As Long, conRow As Long, _
                       conTop As Long, conLeft As Long, _
                       conHeight As Long, conWidth As Long, _
                       conTextAlign As Long, conBackColor As Long)
    
    '/*
    '* コントロール数値設定
    
    With Me.Controls.Add("Forms.Label.1")
        
        .Name = conName & conColumn & conRow: .Caption = conCaption
        
        .Top = conTop: .Left = conLeft
        
        .Height = conHeight: .Width = conWidth
        
        .Font.Name = "メイリオ": .FontSize = 26
        
        .TextAlign = conTextAlign: .BackColor = conBackColor
        
        .SpecialEffect = 3: .BorderStyle = 1
        
    End With

End Sub

Private Sub setControls()
    
    '/*
    '* Class設定
    
    Dim con As Control
    
    For Each con In Me.Controls
    
        If InStr(con.Name, "Label") <> 0 Then

            With New CF_LabelController
            
               ControlCollection.Add .setControlClass(con)
                
            End With
            
        End If

    Next con
    
End Sub

Private Function captionList(n As Long) As String
    
    Select Case n
    
        Case 1
    
            captionList = "C"
            
        Case 2
        
            captionList = "D"
            
        Case 3
            
            captionList = "/"
            
        Case 4
        
            captionList = "*"
            
        Case 5
            
            captionList = "-"
            
        Case 6
        
            captionList = "+"
            
        Case 7
        
            captionList = "="
            
        Case Else
        
            captionList = Str(n)
            
    End Select

End Function

Private Sub R_Click()

    '/*
    '* Form再表示

    Dim i As Long
    For i = 1 To Me.Controls.Count - 1
        Me.Controls.Remove (1)
    Next
    Call UserForm_initialize
    
End Sub
