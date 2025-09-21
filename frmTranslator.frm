VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTranslator 
   Caption         =   "GER-ENG (for ENG-GER place ""-"" at start)"
   ClientHeight    =   3840
   ClientLeft      =   950
   ClientTop       =   3670
   ClientWidth     =   4160
   OleObjectBlob   =   "frmTranslator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'important public variable - look passTextEng
Dim textEng As String
Dim textGer As String
Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub
Sub passText(ByRef textInput As String, ByRef textOutput As String)
'get answer from the Module so its available as a variable in the form
    
    textGer = textInput
    textEng = textOutput
    
End Sub
Private Sub cmdOk_Click()
    
    If Me.txtTextGer = "" Then
               
        If textEng = "" Then
            Me.txtTextGer.Value = "XXXXXXXXXXXX"
        Else
            Me.txtTextGer.Value = textGer
            chooseInput
        End If
        chooseInput
        
    Else
        Unload Me
        Call ModuleTranslator.translator(Me.txtTextGer.Value)
    End If
    
End Sub
Private Sub UserForm_Activate()
    
    frmTranslator.Height = 220
    If textEng = "" Then
        setWidth (160)
    Else
        Dim maxLen As Integer
        maxLen = Application.Large(Array(160, (Len(textEng) + 4) * 6, (Len(textGer) + 4) * 6), 1)
        setWidth (maxLen)
    End If
    
    chooseInput
    
    If textEng = "" Then
        Me.lblPrev.Caption = ""
    Else
        Me.lblPrev.Caption = textEng
        Me.txtTextGer.Value = textGer
        chooseInput
    End If

End Sub
Private Sub UserForm_Click()
    
    chooseInput
    
End Sub
Private Sub setWidth(ByVal wid As Integer)
    
    txtTextGer.Width = wid
    lblPrev.Width = wid
    frmTranslator.Width = wid + 60
    
End Sub
Private Sub chooseInput()
'set focus on text

    'setFocus 2 times - otherwise wont work
    Me.cmdOk.SetFocus
    With Me.txtTextGer
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.text)
    End With

End Sub
