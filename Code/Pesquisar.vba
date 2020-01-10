Option Compare Database
Option Explicit

Private Sub Form_Load()

    Filtro strTabela
    Me.KeyPreview = True
    Me.lstCadastro.SetFocus
    Me.lstCadastro.Selected(1) = True

End Sub

Private Sub cmdFiltrar_Click()

    Dim txtFiltro As String
    txtFiltro = InputBox("Digite uma palavra para fazer o filtro:", "Filtro", , 0, 0)
    Filtro strTabela, txtFiltro
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyInsert
        
           cmdNovo_Click
           
        Case vbKeyReturn
        
            cmdAlterar_Click
            
        Case vbKeyDelete
           
            cmdExcluir_Click
        
        Case vbKeyF2
        
            cmdFiltrar_Click
            
    End Select
End Sub

Private Sub cmdNovo_Click()

    Manipulacao strTabela, "Novo"
    
End Sub

Private Sub cmdAlterar_Click()

    Manipulacao strTabela, "Alterar"
    
End Sub

Private Sub cmdExcluir_Click()

    Manipulacao strTabela, "Excluir"
    
End Sub

Private Sub lstCadastro_DblClick(Cancel As Integer)

    cmdAlterar_Click
    
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Public Sub Manipulacao(Tabela As String, Operacao As String)

If IsNull(Form_Pesquisar.lstCadastro.Value) And Operacao <> "Novo" Then
   Exit Sub
End If

Dim rstFormularios As DAO.Recordset

Set rstFormularios = _
    CurrentDb.OpenRecordset("Select * from Formularios " & _
                            " where TabelaPrincipal = '" & _
                            Tabela & "'")

Select Case Operacao

 Case "Novo"
        
    DoCmd.OpenForm rstFormularios.Fields("NomeDoFormulario"), , , , acFormAdd
    
 Case "Alterar"

    DoCmd.OpenForm rstFormularios.Fields("NomeDoFormulario"), , , rstFormularios.Fields("Identificacao") & " = " & Form_Pesquisar.lstCadastro.Value

 Case "Excluir"

    If MsgBox("Deseja excluir este registro?", vbInformation + vbOKCancel) = vbOK Then
       DoCmd.SetWarnings False
       DoCmd.RunSQL ("Delete from " & strTabela & " where " & rstFormularios.Fields("Identificacao") & " = " & Form_Pesquisar.lstCadastro.Value)
       DoCmd.SetWarnings True
    End If

End Select

Form_Pesquisar.lstCadastro.Requery

Saida:
End Sub

Private Function Filtro(strTabela As String, Optional Procurar As String)

Dim rstFormularios As DAO.Recordset
Dim rstForm_Campos As DAO.Recordset
Dim rstForm_TabRelacionada As DAO.Recordset

Dim Sql As String
Dim SqlAux As String
Dim a As Integer
Dim Colunas As Integer

Set rstFormularios = _
    CurrentDb.OpenRecordset("Select * from Formularios " & _
                            " where TabelaPrincipal = '" & _
                            strTabela & "'")
                                             
Set rstForm_Campos = _
    CurrentDb.OpenRecordset("Select * from Formularios_Campos " & _
                            " where codFormulario = " & _
                            rstFormularios.Fields("codFormulario"))

Set rstForm_TabRelacionada = _
    CurrentDb.OpenRecordset("Select * from Formularios_TabelaRelacionada " & _
                            " where codFormulario = " & _
                            rstFormularios.Fields("codFormulario"))



Sql = "Select "

While Not rstForm_Campos.EOF
    If rstForm_Campos.Fields("Pesquisa") = True Then
        Sql = Sql & IIf(IsNull(rstForm_Campos.Fields("Nome")), _
                      rstForm_Campos.Fields("Campo"), _
                      rstForm_Campos.Fields("Campo") & _
                      " AS " & rstForm_Campos.Fields("Nome")) & ", "
    End If
    
    rstForm_Campos.MoveNext
Wend

Sql = Left(Sql, Len(Sql) - 2) & " "

Sql = Sql & " from " & rstFormularios.Fields("TabelaPrincipal") & ", "
   
Sql = Left(Sql, Len(Sql) - 2) & " "

If Not rstForm_TabRelacionada.EOF Then

    Sql = Sql & ", "

    While Not rstForm_TabRelacionada.EOF
      Sql = Sql & rstForm_TabRelacionada.Fields("TabelaRelacionada") & ", "
      rstForm_TabRelacionada.MoveNext
    Wend
    
    Sql = Left(Sql, Len(Sql) - 2) & " "
    
    SqlAux = ""

    rstForm_TabRelacionada.MoveFirst
    
    While Not rstForm_TabRelacionada.EOF
    
      SqlAux = SqlAux & rstForm_TabRelacionada.Fields("TabelaRelacionada") & "." _
                      & rstForm_TabRelacionada.Fields("CampoChave_Filho") & " = " _
                      & rstFormularios.Fields("TabelaPrincipal") & "." _
                      & rstForm_TabRelacionada.Fields("CampoChave_Pai") & " AND "
                      
      rstForm_TabRelacionada.MoveNext
      
    Wend
    
    If SqlAux <> "" Then
       Sql = Sql & " Where (" & SqlAux
       Sql = Left(Sql, Len(Sql) - 5) & ") "
    End If
    
End If

If SqlAux = "" Then
   Sql = Sql & " Where ("
Else
   Sql = Sql & " AND ("
End If

rstForm_Campos.MoveFirst

While Not rstForm_Campos.EOF
  If rstForm_Campos.Fields("Filtro") = True Then
     Sql = Sql & rstForm_Campos.Fields("Campo") & " Like '*" _
               & LCase(Trim(Procurar)) & "*' OR "
  End If
  rstForm_Campos.MoveNext
Wend

Sql = Left(Sql, Len(Sql) - 3) & ") "

Sql = Sql & "Order By "

rstForm_Campos.MoveFirst

While Not rstForm_Campos.EOF

  If rstForm_Campos.Fields("Ordem") <> "" Then
     Sql = Sql & rstForm_Campos.Fields("Campo") _
               & " " & rstForm_Campos.Fields("Ordem") & ", "
  End If
  
  rstForm_Campos.MoveNext
  
Wend

Sql = Left(Sql, Len(Sql) - 2) & " "

Sql = Sql & ";"

'MsgBox Sql

Me.lstCadastro.RowSource = Sql
Me.lstCadastro.ColumnHeads = True
Me.lstCadastro.ColumnCount = rstForm_Campos.RecordCount
Me.Caption = rstFormularios.Fields("TituloDoFormulario")
'Form_Pesquisar.lstCadastro.ColumnWidths = "0cm;"
'Form_Pesquisar.lstCadastro.ColumnWidths = ""

Dim strTamanho As String

rstForm_Campos.MoveFirst
While Not rstForm_Campos.EOF
  If Not IsNull(rstForm_Campos.Fields("Tamanho")) Then
     strTamanho = strTamanho & str(rstForm_Campos.Fields("Tamanho")) & "cm;"
  End If
  rstForm_Campos.MoveNext
Wend

Me.lstCadastro.ColumnWidths = strTamanho

rstFormularios.Close
rstForm_Campos.Close
rstForm_TabRelacionada.Close

End Function

