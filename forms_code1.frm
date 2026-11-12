VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Documentos Pendentes"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
   OleObjectBlob   =   "Formulário para preenchimento de planilha Excel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CheckBox2_Click()

End Sub

Private Sub CheckBox5_Click()

    txtProgressao.Enabled = CheckBox5.Value

    If Not CheckBox5.Value Then
        txtProgressao.Value = ""
    End If

End Sub


Private Sub cmdLimparTodos_Click()

    CheckBox1.Value = False
    CheckBox2.Value = False
    CheckBox3.Value = False
    CheckBox4.Value = False
    CheckBox5.Value = False
    CheckBox6.Value = False
    CheckBox7.Value = False
    CheckBox8.Value = False

End Sub

Private Sub cmdMarcarTodos_Click()

    CheckBox1.Value = True
    CheckBox2.Value = True
    CheckBox3.Value = True
    CheckBox4.Value = True
    CheckBox5.Value = True
    CheckBox6.Value = True
    CheckBox7.Value = True
    CheckBox8.Value = True

End Sub

Private Sub CommandButton1_Click()

    Dim resultado As String

    If CheckBox1.Value Then resultado = resultado & "• Histórico E. Fundamental" & vbLf
    If CheckBox2.Value Then resultado = resultado & "• Histórico E. Médio" & vbLf
    If CheckBox3.Value Then resultado = resultado & "• Título de Eleitor" & vbLf
    If CheckBox4.Value Then resultado = resultado & "• Certidão de Nascimento ou Casamento" & vbLf
    If CheckBox5.Value Then

    If Trim(txtProgressao.Value) = "" Then
        MsgBox "Você marcou 'Prog. Parcial', mas não informou a matéria e a série.", vbExclamation, "Atenção"
        Exit Sub
    End If

    resultado = resultado & "• Prog. Parcial: " & Trim(txtProgressao.Value) & vbLf

End If
    If CheckBox6.Value Then resultado = resultado & "• Comprovante de Endereço" & vbLf
    If CheckBox7.Value Then resultado = resultado & "• RG e/ou CNH" & vbLf
    If CheckBox8.Value Then resultado = resultado & "• Classificação" & vbLf

    If resultado = "" Then
    resultado = "OK"
    ActiveCell.Font.Color = RGB(0, 0, 0)
    Else
    resultado = Left(resultado, Len(resultado) - 1)
    End If

    ActiveCell.Value = resultado
    ActiveCell.WrapText = True

Unload Me

End Sub

Private Sub CommandButton2_Click()

    Unload Me

End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblAluno_Click()

End Sub

Private Sub txtProgressao_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    CheckBox1.Value = False
    CheckBox2.Value = False
    CheckBox3.Value = False
    CheckBox4.Value = False
    CheckBox5.Value = False
    CheckBox6.Value = False
    CheckBox7.Value = False
    CheckBox8.Value = False

txtProgressao.Value = ""
txtProgressao.Enabled = False

End Sub
