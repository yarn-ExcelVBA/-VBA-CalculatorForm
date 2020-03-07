VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CF_Calclation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private buf() As String
Private TypeCheckFlag As Boolean
Private PrevCheckFlag As Boolean

Public c As Collection

Public Sub Calclate(statusStr As String)
    
    statusStr = StatusValidationCheck(getCalcStatus(statusStr))
    
    If InStr(statusStr, "=") <> 0 Then
    
        CalclaterForm("CalcStatus00").Caption = eval(statusStr)
        
    Else
    
        CalclaterForm("CalcStatus00").Caption = statusStr
        
    End If
    
End Sub

Private Function eval(statusStr As String) As Double

    buf = Split(statusStr, "=")
    
    If InStr(buf(0), "*") <> 0 Then
        
        eval = eval(Mid(buf(0), 1, InStr(buf(0), "*") - 1)) * _
               eval(Mid(buf(0), InStr(buf(0), "*") + 1, Len(buf(0))))
               
    ElseIf InStr(buf(0), "/") <> 0 Then
        
        eval = eval(Mid(buf(0), 1, InStr(buf(0), "/") - 1)) / _
               eval(Mid(buf(0), InStr(buf(0), "/") + 1, Len(buf(0))))
               
    ElseIf InStr(buf(0), "+") <> 0 Then
        
        eval = eval(Mid(buf(0), 1, InStr(buf(0), "+") - 1)) + _
               eval(Mid(buf(0), InStr(buf(0), "+") + 1, Len(buf(0))))
               
    ElseIf InStr(buf(0), "-") <> 0 Then
        
        eval = eval(Mid(buf(0), 1, InStr(buf(0), "-") - 1)) - _
               eval(Mid(buf(0), InStr(buf(0), "-") + 1, Len(buf(0))))
               
    ElseIf InStr(statusStr, "=") = 0 Then
    
        eval = statusStr
        
    Else
        
        eval = buf(0)
               
    End If
    
End Function

Private Property Get StatusValidationCheck(statusStr As String) As String

    Dim i As Long
    
    ReDim strStore(Len(statusStr)) As String
    
    TypeCheckFlag = False
    
    PrevCheckFlag = False
    
    For i = 1 To Len(statusStr)
        
        strStore(i) = Mid(statusStr, i, 1)
        
        Select Case Mid(statusStr, i, 1)
                            
            Case "C"
                StatusValidationCheck = ""
                Exit Sub

            Case "D"
                If statusStr = "D" Then Exit Sub
                
                StatusValidationCheck = Mid(statusStr, 1, Len(statusStr) - 2)
                Exit Sub
                
            Case "="
                If statusStr = "=" Then Exit Sub
                If PrevCheckFlag = True Then
                    StatusValidationCheck = Mid(statusStr, 1, (Len(statusStr) - 1))
                    Exit Sub
                End If
                
            Case "+"
                If statusStr = "+" Then Exit Sub
                If TypeCheckFlag = True Then
                    StatusValidationCheck = Mid(statusStr, 1, Len(statusStr) - 1)
                    Exit Sub
                End If
                
                TypeCheckFlag = True
                PrevCheckFlag = True
                
            Case "-"
                If statusStr = "-" Then Exit Sub
                If TypeCheckFlag = True Then
                    StatusValidationCheck = Mid(statusStr, 1, Len(statusStr) - 1)
                    Exit Sub
                End If
                
                TypeCheckFlag = True
                PrevCheckFlag = True
                
            Case "*"
                If statusStr = "*" Then Exit Sub
                If TypeCheckFlag = True Then
                    StatusValidationCheck = Mid(statusStr, 1, Len(statusStr) - 1)
                    Exit Sub
                End If
                
                TypeCheckFlag = True
                PrevCheckFlag = True
                
            Case "/"
                If statusStr = "/" Then Exit Sub
                If TypeCheckFlag = True Then
                    StatusValidationCheck = Mid(statusStr, 1, Len(statusStr) - 1)
                    Exit Sub
                End If
                
                TypeCheckFlag = True
                PrevCheckFlag = True
                
            Case "."
                If statusStr = "." Then Exit Sub
                If TypeCheckFlag = True Then
                    StatusValidationCheck = Mid(statusStr, 1, Len(statusStr) - 1)
                    Exit Sub
                End If
                
                TypeCheckFlag = True
                PrevCheckFlag = True
                
            Case Else
            
                If InStr(statusStr, "0") = 1 And Len(statusStr) >= 2 Then
                    StatusValidationCheck = Mid(statusStr, 2, (Len(statusStr) - 1))
                    Exit Sub
                End If
            
                PrevCheckFlag = False
                
        End Select
        
    Next i

    StatusValidationCheck = statusStr
    
End Property
    
Private Property Get getCalcStatus(conStr As String) As String

    getCalcStatus = CalclaterForm("CalcStatus00").Caption & conStr
    
End Property

Private Sub calclationSave()

End Sub


