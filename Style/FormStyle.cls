Private clsForm As UserForm
Public bgColor As Long
Public fontColor As Long
Public borderColor As Long
Public fontStyleName As String
Public fontSize As Long

Private Sub Class_Initialize()

    With Me
        .bgColor = RGB(40, 40, 40)
        .borderColor = RGB(40, 40, 40)
        .fontColor = RGB(250, 250, 250)
        .fontStyleName = "Meiryo UI"
        .fontSize = 11
    End With

End Sub

Public Function SetForm(ByRef Form As UserForm) As Boolean

    On Error Resume Next
    
    Set clsForm = Form
    
    On Error GoTo 0
    SetForm = Err.Number
    
End Function

Public Function Style(Optional ByVal name As String = "")

    Select Case LCase(name)
    Case "black"
        With Me
            .bgColor = RGB(40, 40, 40)
            .borderColor = RGB(40, 40, 40)
            .fontColor = RGB(250, 250, 250)
            .fontStyleName = "Meiryo UI"
            .fontSize = 11
        End With
        StyleChangeForm
    Case "blue"
        With Me
            .bgColor = RGB(240, 248, 255)
            .borderColor = RGB(40, 40, 40)
            .fontColor = RGB(0, 0, 0)
            .fontStyleName = "Meiryo UI"
            .fontSize = 11
        End With
        StyleChangeForm
    Case "orange"
        With Me
            .bgColor = RGB(255, 248, 240)
            .borderColor = RGB(40, 40, 40)
            .fontColor = RGB(0, 0, 0)
            .fontStyleName = "Meiryo UI"
            .fontSize = 11
        End With
        StyleChangeForm
    Case "red"
        With Me
            .bgColor = RGB(255, 240, 240)
            .borderColor = RGB(40, 40, 40)
            .fontColor = RGB(0, 0, 0)
            .fontStyleName = "Meiryo UI"
            .fontSize = 11
        End With
        StyleChangeForm
    Case Else

    End Select

End Function

Private Sub StyleChangeForm()
    
    On Error Resume Next
    
    clsForm.SpecialEffect = fmSpecialEffectFlat
    clsForm.BackColor = Me.bgColor
    clsForm.borderColor = Me.borderColor
    clsForm.ForeColor = Me.fontColor
    clsForm.Font.name = Me.fontStyleName
    clsForm.Font.Size = Me.fontSize
    clsForm.Font.Bold = True
    ' https://www.petitmonte.com/bbs/answers?question_id=19864

    For Each ctl In clsForm.Controls
        If TypeName(ctl.Parent) <> "Page" Then
            Select Case TypeName(ctl)
            Case "CommandButton"
                ctl.BackColor = Me.bgColor
                ctl.ForeColor = Me.fontColor
                ctl.Font.name = Me.fontStyleName
                ctl.Font.Size = Me.fontSize
                ctl.Font.Bold = True
            Case "Label"
                ctl.BackColor = Me.bgColor
                ctl.borderColor = Me.bgColor
                ctl.ForeColor = Me.fontColor
                ctl.Font.name = Me.fontStyleName
                ctl.Font.Size = Me.fontSize
                ctl.Font.Bold = True
            Case "MultiPage"
                ctl.BackColor = Me.bgColor
                ctl.ForeColor = RGB(0, 0, 0)
                ctl.Font.name = Me.fontStyleName
                ctl.Font.Size = Me.fontSize
                ctl.Font.Bold = True
                For Each Page In ctl.Pages
                    ' Debug.Print Page.Caption & "[" & TypeName(Page) & "]"
                    Page.BackColor = Me.bgColor
                    Page.ForeColor = Me.fontColor
                Next
            Case Else
                ctl.BackColor = Me.bgColor
                ctl.borderColor = Me.bgColor
                ctl.ForeColor = Me.fontColor
                ctl.Font.name = Me.fontStyleName
                ctl.Font.Size = Me.fontSize
                ctl.Font.Bold = True
                ' Debug.Print ctl.Caption & "[" & TypeName(ctl) & "]"
            End Select
        End If
    Next
    On Error GoTo 0

End Sub


Private Sub Class_Terminate()
    Set clsForm = Nothing
End Sub
