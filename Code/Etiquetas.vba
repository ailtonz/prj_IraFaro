Option Compare Database

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Restore
    Me.Move 200, 600, 11000, 6000
    
End Sub

Private Sub Nome_Click()
    
    Me.Endereco = Me.Nome.Column(1)
    Me.Bairro = Me.Nome.Column(2)
    Me.Cep = Me.Nome.Column(3)
    Me.Municipio = Me.Nome.Column(4)
    Me.Estado = Me.Nome.Column(5)
    
End Sub
