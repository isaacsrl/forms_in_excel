Attribute VB_Name = "Módulo1"
Sub AbrirDocumentos()
Attribute AbrirDocumentos.VB_Description = "Documentos Pendentes"
Attribute AbrirDocumentos.VB_ProcData.VB_Invoke_Func = "D\n14"

    If ActiveCell.Column <> 3 Then
        MsgBox "Selecione primeiro a célula de Documentos Faltantes do aluno.", vbExclamation, "Atenção"
        Exit Sub
    End If

    If ActiveCell.Row < 3 Then
        MsgBox "Selecione uma célula de um aluno.", vbExclamation, "Atenção"
        Exit Sub
    End If

    UserForm1.lblAluno.Caption = "ALUNO: " & Cells(ActiveCell.Row, 2).Value

    UserForm1.Show

End Sub
