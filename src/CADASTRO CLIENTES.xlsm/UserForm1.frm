VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   12360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub btAddVeiculo_Click()

'
'    If lblRenavam2.Visible = False Then ' Verifica se não está visível
'        lblRenavam2.Visible = True ' Mostra o label
'        txtRenavam2.Visible = True ' Mostra o campo de texto
'        lblPlaca2.Visible = True ' Mostra o label
'        txtPlaca2.Visible = True ' Mostra o campo de texto
'        btAddVeiculo.Caption = "-" ' muda para o sinal de - para remover a exibição novamente
'    Else
'        lblRenavam2.Visible = False ' Oculta label
'        txtRenavam2.Visible = False ' Oculta o campo de texto
'        lblPlaca2.Visible = False ' Oculta o label
'        txtPlaca2.Visible = False ' Oculta o campo de texto
'        btAddVeiculo.Caption = "+" ' muda para o sinal de + para poder exibir novamente
'    End If
'
'End Sub

Private contadorVeiculo As Integer ' Variável para controlar o número de campos adicionados
Private contadorCassacao As Integer
Private contadorSuspensao As Integer

Private Sub btAddPaCassacao_Click()
    If contadorCassacao < 2 Then ' Verifica se o limite de campos foi atingido (2 campos)
        contadorCassacao = contadorCassacao + 1 ' Incrementa o contador
        
        ' Calcula o índice dos campos adicionados
        Dim indiceCassacao As String
        indiceCassacao = CStr(contadorCassacao + 1)
        
        ' Mostra os controles correspondentes ao campo adicionado
        Me.Controls("lblPaCassacao" & indiceCassacao).Visible = True
        Me.Controls("txtPaCassacao" & indiceCassacao).Visible = True
        
        btRemovePaCassacao.Enabled = True ' Habilita o botão de remoção
        
        If contadorCassacao = 2 Then ' Se o limite de campos for atingido, desabilita o botão de adição
            btAddPaCassacao.Enabled = False
        End If
    End If
End Sub

Private Sub btnAbrirImagem_Click()
    ' Abre a caixa de diálogo para selecionar a imagem
    Dim filePath As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "Selecione uma imagem"
        .Filters.Add "Imagens", "*.jpg;*.jpeg;*.png;*.gif", 1
        .AllowMultiSelect = False
        
        If .Show = -1 Then ' Clique em "Abrir" na caixa de diálogo
            filePath = .SelectedItems(1)
            
            ' Carrega e exibe a imagem no controle Image
            Me.Image1.Picture = LoadPicture(filePath)
        End If
    End With
    
    Set fileDialog = Nothing
End Sub



Private Sub btPesquisar_Click()

'Dim Y As Long
'Dim ultimodadoNome As Long
'Dim ultimodadoCpf As Long
'Dim linha As Long
'
'If Me.txtPesquisarNome.Value = "" Then
'MsgBox "Digite um Nome para buscar!"
'Exit Sub
'End If
'
'If Me.txtPesquisarCpf.Value = "" Then
'MsgBox "Digite um Cpf para buscar!"
'Exit Sub
'End If
'
'ultimodadoNome = Planilha1.Range("B" & Rows.Count).End(xlUp).Row
'ultimodadoCpf = Planilha1.Range("D" & Rows.Count).End(xlUp).Row
'
'Lista.RowSource = Null
'
'Y = 0
'
'For linha = 4 To ultimodadoNome
'Nome = ActiveSheet.Cell(linha, 2).Value
'
'    If Nome Like Me.txtPesquisarNome.Value Then
'        Me.Lista.AddItem
'
'
'        Me.Lista.List(Y, 0) = ActiveSheet.Cells(linha, 1).Value
'        Me.Lista.List(Y, 1) = ActiveSheet.Cells(linha, 2).Value
'        Me.Lista.List(Y, 2) = ActiveSheet.Cells(linha, 3).Value
'        Me.Lista.List(Y, 3) = ActiveSheet.Cells(linha, 4).Value
'        Me.Lista.List(Y, 4) = ActiveSheet.Cells(linha, 5).Value
'        Me.Lista.List(Y, 5) = ActiveSheet.Cells(linha, 6).Value
'        Me.Lista.List(Y, 6) = ActiveSheet.Cells(linha, 7).Value
'        Me.Lista.List(Y, 7) = ActiveSheet.Cells(linha, 8).Value
'        Me.Lista.List(Y, 8) = ActiveSheet.Cells(linha, 9).Value
'        Me.Lista.List(Y, 9) = ActiveSheet.Cells(linha, 10).Value
'        Me.Lista.List(Y, 10) = ActiveSheet.Cells(linha, 11).Value
'        Me.Lista.List(Y, 11) = ActiveSheet.Cells(linha, 12).Value
'        Me.Lista.List(Y, 12) = ActiveSheet.Cells(linha, 13).Value
'        Me.Lista.List(Y, 13) = ActiveSheet.Cells(linha, 14).Value
'        Me.Lista.List(Y, 14) = ActiveSheet.Cells(linha, 15).Value
'        Me.Lista.List(Y, 15) = ActiveSheet.Cells(linha, 16).Value
'        Me.Lista.List(Y, 16) = ActiveSheet.Cells(linha, 17).Value
'        Me.Lista.List(Y, 17) = ActiveSheet.Cells(linha, 18).Value
'        Me.Lista.List(Y, 18) = ActiveSheet.Cells(linha, 19).Value
'        Me.Lista.List(Y, 19) = ActiveSheet.Cells(linha, 20).Value
'        Me.Lista.List(Y, 20) = ActiveSheet.Cells(linha, 21).Value
'        Me.Lista.List(Y, 21) = ActiveSheet.Cells(linha, 22).Value
'        Me.Lista.List(Y, 22) = ActiveSheet.Cells(linha, 23).Value
'        Me.Lista.List(Y, 23) = ActiveSheet.Cells(linha, 24).Value
'        Me.Lista.List(Y, 24) = ActiveSheet.Cells(linha, 25).Value
'        Me.Lista.List(Y, 25) = ActiveSheet.Cells(linha, 26).Value
'        Me.Lista.List(Y, 26) = ActiveSheet.Cells(linha, 27).Value
'        Me.Lista.List(Y, 27) = ActiveSheet.Cells(linha, 28).Value
'        Me.Lista.List(Y, 28) = ActiveSheet.Cells(linha, 29).Value
'        Me.Lista.List(Y, 29) = ActiveSheet.Cells(linha, 30).Value
'        Me.Lista.List(Y, 30) = ActiveSheet.Cells(linha, 31).Value
'        Me.Lista.List(Y, 31) = ActiveSheet.Cells(linha, 32).Value
'        Me.Lista.List(Y, 32) = ActiveSheet.Cells(linha, 33).Value
'        Me.Lista.List(Y, 33) = ActiveSheet.Cells(linha, 34).Value
'        Me.Lista.List(Y, 34) = ActiveSheet.Cells(linha, 35).Value
'        Me.Lista.List(Y, 35) = ActiveSheet.Cells(linha, 36).Value
'        Me.Lista.List(Y, 36) = ActiveSheet.Cells(linha, 37).Value
'        Me.Lista.List(Y, 37) = ActiveSheet.Cells(linha, 38).Value
'        Me.Lista.List(Y, 38) = ActiveSheet.Cells(linha, 39).Value
'        Me.Lista.List(Y, 39) = ActiveSheet.Cells(linha, 40).Value
'        Me.Lista.List(Y, 40) = ActiveSheet.Cells(linha, 41).Value
'        Me.Lista.List(Y, 41) = ActiveSheet.Cells(linha, 42).Value
'        Me.Lista.List(Y, 42) = ActiveSheet.Cells(linha, 43).Value
'        Me.Lista.List(Y, 43) = ActiveSheet.Cells(linha, 44).Value
'        Me.Lista.List(Y, 44) = ActiveSheet.Cells(linha, 45).Value
'        Me.Lista.List(Y, 45) = ActiveSheet.Cells(linha, 46).Value
'        Me.Lista.List(Y, 46) = ActiveSheet.Cells(linha, 47).Value
'        Me.Lista.List(Y, 47) = ActiveSheet.Cells(linha, 48).Value
'        Me.Lista.List(Y, 48) = ActiveSheet.Cells(linha, 49).Value
'        Me.Lista.List(Y, 49) = ActiveSheet.Cells(linha, 50).Value
'        Me.Lista.List(Y, 50) = ActiveSheet.Cells(linha, 51).Value
'        Me.Lista.List(Y, 51) = ActiveSheet.Cells(linha, 52).Value
'        Me.Lista.List(Y, 52) = ActiveSheet.Cells(linha, 53).Value
'        Me.Lista.List(Y, 53) = ActiveSheet.Cells(linha, 54).Value
'
'        Y = Y + 1
'
'        End If
'
'            Next
        
End Sub

Private Sub btRemovePaCassacao_Click()

    If contadorCassacao > 0 Then ' Verifica se existem campos adicionados
        ' Calcula o índice dos campos a serem removidos
        Dim indiceCassacao As String
        indiceCassacao = CStr(contadorCassacao + 1)
        
        ' Oculta os controles correspondentes ao campo a ser removido
        Me.Controls("lblPaCassacao" & indiceCassacao).Visible = False
        Me.Controls("txtPaCassacao" & indiceCassacao).Visible = False
        
        contadorCassacao = contadorCassacao - 1 ' Decrementa o contador
        
        If contadorCassacao = 0 Then ' Se não houver mais campos adicionados, desabilita o botão de remoção
            btRemovePaCassacao.Enabled = False
        End If
        
        btAddPaCassacao.Enabled = True ' Habilita o botão de adição
        
        btAddPaCassacao.Caption = "+" ' Restaura o texto do botão "Add"
    End If
End Sub


Private Sub btAddPaSuspensao_Click()
    If contadorSuspensao < 2 Then ' Verifica se o limite de campos foi atingido (2 campos)
        contadorSuspensao = contadorSuspensao + 1 ' Incrementa o contador
        
        ' Calcula o índice dos campos adicionados
        Dim indiceSuspensao As String
        indiceSuspensao = CStr(contadorSuspensao + 1)
        
        ' Mostra os controles correspondentes ao campo adicionado
        Me.Controls("lblPaSuspensao" & indiceSuspensao).Visible = True
        Me.Controls("txtPaSuspensao" & indiceSuspensao).Visible = True
        
        btRemovePaSuspensao.Enabled = True ' Habilita o botão de remoção
        
        If contadorSuspensao = 2 Then ' Se o limite de campos for atingido, desabilita o botão de adição
            btAddPaSuspensao.Enabled = False
        End If
    End If
End Sub


Private Sub btRemovePaSuspensao_Click()
    If contadorSuspensao > 0 Then ' Verifica se existem campos adicionados
        ' Calcula o índice dos campos a serem removidos
        Dim indiceSuspensao As String
        indiceSuspensao = CStr(contadorSuspensao + 1)
        
        ' Oculta os controles correspondentes ao campo a ser removido
        Me.Controls("lblPaSuspensao" & indiceSuspensao).Visible = False
        Me.Controls("txtPaSuspensao" & indiceSuspensao).Visible = False
        
        contadorSuspensao = contadorSuspensao - 1 ' Decrementa o contador
        
        If contadorSuspensao = 0 Then ' Se não houver mais campos adicionados, desabilita o botão de remoção
            btRemovePaSuspensao.Enabled = False
        End If
        
        btAddPaSuspensao.Enabled = True ' Habilita o botão de adição
        
        btAddPaSuspensao.Caption = "+" ' Restaura o texto do botão "Add"
    End If
End Sub

Private Sub btAddVeiculo_Click()
    If contadorVeiculo < 2 Then ' Verifica se o limite de campos foi atingido (2 campos)
        contadorVeiculo = contadorVeiculo + 1 ' Incrementa o contador
        
        ' Calcula o índice dos campos adicionados
        Dim indiceVeiculo As String
        indiceVeiculo = CStr(contadorVeiculo + 1)
        
        ' Mostra os controles correspondentes ao campo adicionado
        Me.Controls("lblRenavam" & indiceVeiculo).Visible = True
        Me.Controls("txtRenavam" & indiceVeiculo).Visible = True
        Me.Controls("lblPlaca" & indiceVeiculo).Visible = True
        Me.Controls("txtPlaca" & indiceVeiculo).Visible = True
        
        btRemoveVeiculo.Enabled = True ' Habilita o botão de remoção
        
        If contadorVeiculo = 2 Then ' Se o limite de campos for atingido, desabilita o botão de adição
            btAddVeiculo.Enabled = False
        End If
    End If
End Sub


Private Sub btRemoveVeiculo_Click()
    If contadorVeiculo > 0 Then ' Verifica se existem campos adicionados
        ' Calcula o índice dos campos a serem removidos
        Dim indiceVeiculo As String
        indiceVeiculo = CStr(contadorVeiculo + 1)
        
        ' Oculta os controles correspondentes ao campo a ser removido
        Me.Controls("lblRenavam" & indiceVeiculo).Visible = False
        Me.Controls("txtRenavam" & indiceVeiculo).Visible = False
        Me.Controls("lblPlaca" & indiceVeiculo).Visible = False
        Me.Controls("txtPlaca" & indiceVeiculo).Visible = False
        
        contadorVeiculo = contadorVeiculo - 1 ' Decrementa o contador
        
        If contadorVeiculo = 0 Then ' Se não houver mais campos adicionados, desabilita o botão de remoção
            btRemoveVeiculo.Enabled = False
        End If
        
        btAddVeiculo.Enabled = True ' Habilita o botão de adição
        
        btAddVeiculo.Caption = "+" ' Restaura o texto do botão "Add"
    End If
End Sub

Private Sub btNovo_Click()
    
        Dim wsId As Worksheet
        Set wsId = ThisWorkbook.Sheets("Gerar ID")
    
        ' foco no primeiro campo do formulário
        Me.txtNome.SetFocus

         ' Limpa os campos
        Me.txtId.Value = ""
        Me.txtNome.Value = ""
        Me.txtNacionalidade.Value = ""
        Me.txtCpf.Value = ""
        Me.txtRg.Value = ""
        
        Me.txtNumero.Value = ""
        
        Me.txtCep.Value = ""
        Me.txtComplemento.Value = ""
        Me.txtDataNascimento.Value = ""
        Me.cbbEstadoCivil.Value = ""
        Me.txtProfissao.Value = ""
        Me.txtTelefone.Value = ""
        Me.txtTelefone2.Value = ""
        Me.txtTelefoneRecado.Value = ""
        Me.txtEmail.Value = ""
        Me.txtServico.Value = ""
        Me.txtValor.Value = ""
        Me.cbbPagamento.Value = ""
        Me.txtCondicoesParce.Value = ""
        Me.txtCnh.Value = ""
        Me.cbbCategoria.Value = ""
        Me.txtVencimentoCnh.Value = ""
        Me.txtEspelho.Value = ""
        Me.txtRenach.Value = ""
        Me.txtCodSeguranca.Value = ""
        Me.txtRenavam1.Value = ""
        Me.txtPlaca1.Value = ""
        Me.txtRenavam2.Value = ""
        Me.txtPlaca2.Value = ""
        Me.txtRenavam3.Value = ""
        Me.txtPlaca3.Value = ""
        Me.txtLoginDetran.Value = ""
        Me.txtSenhaDetran.Value = ""
        Me.txtLoginGov.Value = ""
        Me.txtSenhaGov.Value = ""
        Me.txtLoginPoupatempo.Value = ""
        Me.txtSenhaPoupatempo.Value = ""
        Me.txtPaCassacao1.Value = ""
        Me.txtPaCassacao2.Value = ""
        Me.txtPaCassacao3.Value = ""
        Me.txtPaSuspensao1.Value = ""
        Me.txtPaSuspensao2.Value = ""
        Me.txtPaSuspensao3.Value = ""
        Me.txtRamo.Value = ""
        Me.txtPessoaResp.Value = ""
        Me.txtNumeroProcesso.Value = ""
        Me.txtForo.Value = ""
        Me.txtVara.Value = ""
        Me.txtDataDistribuicao.Value = ""
        Me.txtAdvogado.Value = ""
        
        ' buscando função de gerar o Id
        Me.txtId.Value = wsId.Cells(2, 1).Value

End Sub

Private Sub btProximo1_Click()

   ' Verifica se há campos obrigatórios vazios
    If CampoVazio(UserForm1) Then
        If Me.txtId.Text = "" Or _
           Me.txtNome.Text = "" Or _
           Me.txtCpf.Text = "" Or _
           Me.txtRg.Text = "" Or _
           Me.txtDigito.Text = "" Or _
           Me.cbbUfRg.Text = "" Or _
           Me.txtNacionalidade.Text = "" Or _
           Me.txtDataNascimento.Text = "" Or _
           Me.cbbEstadoCivil.Text = "" Or _
           Me.txtProfissao.Text = "" Or _
           Me.txtCep.Text = "" Or _
           Me.txtNumero.Text = "" Or _
           Me.txtDdd1.Text = "" Or _
           Me.txtTelefone.Text = "" Or _
           Me.txtEmail.Text = "" Or _
           Me.txtServico.Text = "" Or _
           Me.txtValor.Text = "" Or _
           Me.cbbPagamento.Text = "" Or _
           Me.txtCondicoesParce.Text = "" Then
           

            
        Else
        
            ' MsgBox "Todos os Campos são Obrigatórios!"
            DefinirFocoPrimeiraPagina
            
        End If
        Else
                    Dim i As Integer
            For i = 1 To Me.MultiPage1.Pages.Count - 1
                Me.MultiPage1.Pages(i).Enabled = True
            Next i
            ' Navega para a próxima página
            If Me.MultiPage1.Value < Me.MultiPage1.Pages.Count - 1 Then
                Me.MultiPage1.Value = Me.MultiPage1.Value + 1
            End If
        
    End If

End Sub

Private Sub btProximo2_Click()

    If Me.MultiPage1.Value < Me.MultiPage1.Pages.Count - 1 Then
        Me.MultiPage1.Value = Me.MultiPage1.Value + 1
    End If

End Sub

'Private Sub btRemoveVeiculo_Click()
'
'End Sub

Private Sub btSalvar_Click()
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim wsId As Worksheet
    Set wsId = ThisWorkbook.Sheets("Gerar ID")
    
    ws.Rows(4).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' transf. labels do endereco
    Dim textoLogradouro As String
    textoLogradouro = Me.lbLogradouro.Caption
    
    Dim textoBairro As String
    textoBairro = Me.lbBairro.Caption
    
    Dim textoCidade As String
    textoCidade = Me.lbCidade.Caption
    
    Dim textoEstado As String
    textoEstado = Me.lbEstado.Caption
    
    
    With ws
        .Range("A4").Value = Me.txtId.Value
        .Range("B4").Value = Me.txtNome.Value
        .Range("C4").Value = Me.txtNacionalidade.Value
        .Range("D4").Value = Me.txtCpf.Value

'=====================================================================
        

        ' concatenando RG
        Dim rg As String
        Dim digito As String
        Dim uf As String
     
        rg = Me.txtRg.Value
        digito = Me.txtDigito.Value
        uf = Me.cbbUfRg.Value
         
        rg = rg & "-" & digito & " SSP/" & uf ' Concatenação dos campos
        
        ws.Range("E4").Value = rg
        ' .Range("E4").Value = Me.txtRg.Value

'=====================================================================
         
        .Range("F4").Value = textoLogradouro
        .Range("G4").Value = textoBairro
        .Range("H4").Value = Me.txtNumero.Value
        .Range("I4").Value = textoCidade
        .Range("J4").Value = textoEstado
        .Range("K4").Value = Me.txtCep.Value
        .Range("L4").Value = Me.txtComplemento.Value
        .Range("M4").Value = Me.txtDataNascimento.Value
        .Range("N4").Value = Me.cbbEstadoCivil.Value
        .Range("O4").Value = Me.txtProfissao.Value
        
'=====================================================================

        ' concatenando telefone 1
        Dim ddd1 As String
        Dim tel As String
        Dim telefone As String
     
        ddd1 = Me.txtDdd1.Value
        tel = Me.txtTelefone.Value
         
        telefone = ddd1 & " " & tel ' Concatenação dos campos
        
        ws.Range("P4").Value = telefone
        ' .Range("P4").Value = Me.txtTelefone.Value
        
'=====================================================================
        
        ' concatenando telefone 2
        Dim ddd2 As String
        Dim tel2 As String
        Dim telefone2 As String
     
        ddd2 = Me.txtDdd2.Value
        tel2 = Me.txtTelefone2.Value
         
        telefone2 = ddd2 & " " & tel2 ' Concatenação dos campos
        
        ws.Range("Q4").Value = telefone2
        ' .Range("Q4").Value = Me.txtTelefone2.Value

'=====================================================================
        
        ' concatenando telefone recado
        Dim dddRecado As String
        Dim telRecado As String
        Dim telefoneRecado As String
     
        dddRecado = Me.txtDddRecado.Value
        telRecado = Me.txtTelefoneRecado.Value
         
        telefoneRecado = dddRecado & " " & telRecado ' Concatenação dos campos
        
         ws.Range("R4").Value = telefoneRecado
        ' .Range("R4").Value = Me.txtTelefoneRecado.Value
        
'=====================================================================

        .Range("S4").Value = Me.txtEmail.Value
        
        .Range("T4").Value = Me.txtServico.Value
        .Range("U4").Value = Me.txtValor.Value
        .Range("V4").Value = Me.cbbPagamento.Value
        .Range("W4").Value = Me.txtCondicoesParce.Value
        
        .Range("X4").Value = Me.txtCnh.Value
        .Range("Y4").Value = Me.cbbCategoria.Value
        
        .Range("Z4").Value = Me.txtVencimentoCnh.Value
        .Range("AA4").Value = Me.txtEspelho.Value
        .Range("AB4").Value = Me.txtRenach.Value
        
        .Range("AC4").Value = Me.txtCodSeguranca.Value
        
        .Range("AD4").Value = Me.txtRenavam1.Value
        .Range("AE4").Value = Me.txtPlaca1.Value
        .Range("AF4").Value = Me.txtRenavam2.Value
        .Range("AG4").Value = Me.txtPlaca2.Value
        .Range("AH4").Value = Me.txtRenavam3.Value
        .Range("AI4").Value = Me.txtPlaca3.Value
        
        .Range("AJ4").Value = Me.txtLoginDetran.Value
        .Range("AK4").Value = Me.txtSenhaDetran.Value
        .Range("AL4").Value = Me.txtLoginGov.Value
        .Range("AM4").Value = Me.txtSenhaGov.Value
        
        .Range("AN4").Value = Me.txtLoginPoupatempo.Value
        .Range("AO4").Value = Me.txtSenhaPoupatempo.Value
 
        
        .Range("AP4").Value = Me.txtPaCassacao1.Value
        .Range("AQ4").Value = Me.txtPaCassacao2.Value
        .Range("AR4").Value = Me.txtPaCassacao3.Value
        
        .Range("AS4").Value = Me.txtPaSuspensao1.Value
        .Range("AT4").Value = Me.txtPaSuspensao2.Value
        .Range("AU4").Value = Me.txtPaSuspensao3.Value
        
        
        .Range("AV4").Value = Me.txtRamo.Value
        .Range("AW4").Value = Me.txtPessoaResp.Value
        .Range("AX4").Value = Me.txtNumeroProcesso.Value
        .Range("AY4").Value = Me.txtForo.Value
        .Range("AZ4").Value = Me.txtVara.Value
        .Range("BA4").Value = Me.txtDataDistribuicao.Value
        .Range("BB4").Value = Me.txtAdvogado.Value
        
        

        
        
    End With
    
        ' Limpa os campos
        Me.txtId.Value = ""
        Me.txtNome.Value = ""
        Me.txtNacionalidade.Value = ""
        Me.txtCpf.Value = ""
        Me.txtRg.Value = ""
        
        Me.txtNumero.Value = ""
        
        Me.txtCep.Value = ""
        Me.txtComplemento.Value = ""
        Me.txtDataNascimento.Value = ""
        Me.cbbEstadoCivil.Value = ""
        Me.txtProfissao.Value = ""
        Me.txtTelefone.Value = ""
        Me.txtTelefone2.Value = ""
        Me.txtTelefoneRecado.Value = ""
        Me.txtEmail.Value = ""
        Me.txtServico.Value = ""
        Me.txtValor.Value = ""
        Me.cbbPagamento.Value = ""
        Me.txtCondicoesParce.Value = ""
        Me.txtCnh.Value = ""
        Me.cbbCategoria.Value = ""
        Me.txtVencimentoCnh.Value = ""
        Me.txtEspelho.Value = ""
        Me.txtRenach.Value = ""
        Me.txtCodSeguranca.Value = ""
        Me.txtRenavam1.Value = ""
        Me.txtPlaca1.Value = ""
        Me.txtRenavam2.Value = ""
        Me.txtPlaca2.Value = ""
        Me.txtRenavam3.Value = ""
        Me.txtPlaca3.Value = ""
        Me.txtLoginDetran.Value = ""
        Me.txtSenhaDetran.Value = ""
        Me.txtLoginGov.Value = ""
        Me.txtSenhaGov.Value = ""
        Me.txtLoginPoupatempo.Value = ""
        Me.txtSenhaPoupatempo.Value = ""
        Me.txtPaCassacao1.Value = ""
        Me.txtPaCassacao2.Value = ""
        Me.txtPaCassacao3.Value = ""
        Me.txtPaSuspensao1.Value = ""
        Me.txtPaSuspensao2.Value = ""
        Me.txtPaSuspensao3.Value = ""
        Me.txtRamo.Value = ""
        Me.txtPessoaResp.Value = ""
        Me.txtNumeroProcesso.Value = ""
        Me.txtForo.Value = ""
        Me.txtVara.Value = ""
        Me.txtDataDistribuicao.Value = ""
        Me.txtAdvogado.Value = ""
    
        ' buscando função de gerar o Id
         Me.txtId.Value = wsId.Cells(2, 1).Value
        
         
        MsgBox "Cadastro realizado!"

     
End Sub


Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub CommandButton8_Click()


End Sub

Private Sub Label27_Click()

End Sub

Private Sub btVoltar1_Click()

    If Me.MultiPage1.Value > 0 Then
        Me.MultiPage1.Value = Me.MultiPage1.Value - 1
    End If

End Sub

Private Sub btVoltar2_Click()

    If Me.MultiPage1.Value > 0 Then
        Me.MultiPage1.Value = Me.MultiPage1.Value - 1
    End If

End Sub

Private Sub btVoltar3_Click()

    If Me.MultiPage1.Value > 0 Then
        Me.MultiPage1.Value = Me.MultiPage1.Value - 1
    End If

End Sub

Private Sub cbbCategoria_Change()

End Sub

Public Sub cbbEstadoCivil_Change()

End Sub


Private Sub Frame2_Click()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Endereco_Click()

End Sub

Private Sub Frame6_Click()

End Sub

Private Sub Frame9_Click()

End Sub

Private Sub Label31_Click()


End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox24_Change()

End Sub

Private Sub Label36_Click()

End Sub

Private Sub Label37_Click()

End Sub

Private Sub Label38_Click()

End Sub

Private Sub lbErro_Click()

End Sub

Private Sub lblAdvogado_Click()

End Sub

Private Sub lblNome_Click()

End Sub

Private Sub lbLogradouro_Click()

End Sub

Private Sub lblSenhaDetran_Click()

End Sub

Private Sub lblServico_Click()

End Sub

Private Sub lblTitulo2_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub logo_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub MultiPage1_Change()
'     If MultiPage1.Value > 0 Then  ' Verifica se está tentando acessar a próxima página
'        If Not CamposObrigatoriosPreenchidos() Then ' Verifica se os campos obrigatórios foram preenchidos
'            MsgBox "Preencha todos os campos obrigatórios antes de prosseguir."
'            ' DefinirFocoPrimeiraPagina
'        End If
'    End If
    If MultiPage1.Value = 0 Then ' Verifica se é a primeira página
        Dim camposVazios As Boolean
        camposVazios = CampoVazio(UserForm1) ' Verifica se há campos obrigatórios vazios
        
        ' Habilita as abas das páginas subsequentes se não houver campos obrigatórios vazios
        Me.MultiPage1.Pages(1).Enabled = Not camposVazios
        Me.MultiPage1.Pages(2).Enabled = Not camposVazios
        ' Adicione linhas semelhantes para outras páginas subsequentes, se houver
    End If
    
    
End Sub

Private Function CamposObrigatoriosPreenchidos() As Boolean
    Dim campos As Variant
    campos = Array(Me.txtNome, Me.txtNacionalidade, Me.txtCpf, Me.txtRg, _
                   Me.txtNumero, Me.txtCep, Me.txtComplemento, Me.txtDataNascimento, _
                   Me.cbbEstadoCivil, Me.txtProfissao, Me.txtTelefone, _
                   Me.txtEmail, Me.txtServico, Me.txtValor, _
                   Me.cbbPagamento, Me.txtCondicoesParce)
    
    Dim campo As Variant
    For Each campo In campos
        If campo.Value = "" Then
            CamposObrigatoriosPreenchidos = False ' Retorna falso se algum campo estiver vazio
            Exit Function
        End If
    Next campo
    
    ' Se todos os campos estão preenchidos, habilita as páginas seguintes
    Me.MultiPage1.Pages(1).Enabled = True
    Me.MultiPage1.Pages(2).Enabled = True
    Me.MultiPage1.Pages(3).Enabled = True
    
    CamposObrigatoriosPreenchidos = True ' Retorna verdadeiro se todos os campos estiverem preenchidos
End Function

Private Sub TextBox2_Change()

End Sub

Private Sub txtCep_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Dim cep As String
    Dim resposta As String
    If KeyCode <> 13 Then Exit Sub
    
    cep = txtCep.Value
    
    ' Limpar o conteúdo do Label de erro
    lbErro.Caption = ""
    
    resposta = BuscarCEP(cep)
    
    If resposta = "CEP NÃO ENCONTRADO!" Then
        Call LimparCampos
        lbErro.Caption = "CEP NÃO ENCONTRADO!"
        Exit Sub
    End If
    
    ' UCase deixa todas as letras maiúsculas
    'controles
    lbLogradouro.Caption = UCase(Parse(resposta, 2))
    lbBairro.Caption = UCase(Parse(resposta, 4))
    lbCidade.Caption = UCase(Parse(resposta, 5))
    lbEstado.Caption = UCase(Parse(resposta, 6))
    
End Sub

Private Sub txtCpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.txtCpf.MaxLength = 14
    
    If Len(Me.txtCpf) = 3 Then
    Me.txtCpf.Text = Me.txtCpf.Text & "."
    Me.txtCpf.SelStart = Len(Me.txtCpf)
    End If
    
    If Len(Me.txtCpf) = 7 Then
    Me.txtCpf.Text = Me.txtCpf.Text & "."
    Me.txtCpf.SelStart = Len(Me.txtCpf)
    End If
    
    If Len(Me.txtCpf) = 11 Then
    Me.txtCpf.Text = Me.txtCpf.Text & "."
    Me.txtCpf.SelStart = Len(Me.txtCpf)
    End If


End Sub

Private Sub txtDataDistribuicao_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.txtDataDistribuicao.MaxLength = 10
    
    If Len(Me.txtDataDistribuicao) = 2 Then
    Me.txtDataDistribuicao.Text = Me.txtDataDistribuicao.Text & "/"
    Me.txtDataDistribuicao.SelStart = Len(Me.txtDataDistribuicao)
    End If
    
    If Len(Me.txtDataDistribuicao) = 5 Then
    Me.txtDataDistribuicao.Text = Me.txtDataDistribuicao.Text & "/"
    Me.txtDataDistribuicao.SelStart = Len(Me.txtDataDistribuicao)
    End If

End Sub

Private Sub txtDataNascimento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.txtDataNascimento.MaxLength = 10
    
    If Len(Me.txtDataNascimento) = 2 Then
    Me.txtDataNascimento.Text = Me.txtDataNascimento.Text & "/"
    Me.txtDataNascimento.SelStart = Len(Me.txtDataNascimento)
    End If
    
    If Len(Me.txtDataNascimento) = 5 Then
    Me.txtDataNascimento.Text = Me.txtDataNascimento.Text & "/"
    Me.txtDataNascimento.SelStart = Len(Me.txtDataNascimento)
    End If

End Sub

Private Sub txtDdd1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Me.txtDdd1.MaxLength = 2
    
    If KeyCode = vbKeyReturn Then
        KeyCode = 0 ' Cancela o Enter para não gerar nova linha
     
        ' Verifica se o DDD está formatado corretamente com os parênteses
        If Not DDDFormatado(Me.txtDdd1.Text) Then
            ' Verifica se o DDD possui 2 dígitos ou está vazio
            If Len(Me.txtDdd1) = 2 Or Len(Me.txtDdd1) = 0 Then
                If Len(Me.txtDdd1) = 2 Then
                    Me.txtDdd1.Text = "(" & Me.txtDdd1.Text & ")"
                    Me.txtDdd1.SelStart = Len(Me.txtDdd1)
                End If
                SendKeys "{TAB}"
            Else
                MsgBox "O DDD deve conter dois dígitos!"
            End If
        Else
            SendKeys "{TAB}"
        End If
    End If
End Sub

Function DDDFormatado(ByVal texto As String) As Boolean
    DDDFormatado = False
    If Left(texto, 1) = "(" And Right(texto, 1) = ")" Then
        DDDFormatado = True
    End If
End Function

Private Sub txtDdd2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Me.txtDdd2.MaxLength = 2
    
    If KeyCode = vbKeyReturn Then
        KeyCode = 0 ' Cancela o Enter para não gerar nova linha
        
        ' Verifica se o DDD está formatado corretamente com os parênteses
        If Not DDDFormatado(Me.txtDdd2.Text) Then
            ' Verifica se o DDD possui 2 dígitos ou está vazio
            If Len(Me.txtDdd2) = 2 Or Len(Me.txtDdd2) = 0 Then
                If Len(Me.txtDdd2) = 2 Then
                    Me.txtDdd2.Text = "(" & Me.txtDdd2.Text & ")"
                    Me.txtDdd2.SelStart = Len(Me.txtDdd2)
                End If
                SendKeys "{TAB}"
            Else
                MsgBox "O DDD deve conter dois dígitos!"
            End If
        Else
            SendKeys "{TAB}"
        End If
    End If
    
End Sub

Private Sub txtDddRecado_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Me.txtDddRecado.MaxLength = 2
    
    If KeyCode = vbKeyReturn Then
        KeyCode = 0 ' Cancela o Enter para não gerar nova linha
        
        ' Verifica se o DDD está formatado corretamente com os parênteses
        If Not DDDFormatado(Me.txtDddRecado.Text) Then
            ' Verifica se o DDD possui 2 dígitos ou está vazio
            If Len(Me.txtDddRecado) = 2 Or Len(Me.txtDddRecado) = 0 Then
                If Len(Me.txtDddRecado) = 2 Then
                    Me.txtDddRecado.Text = "(" & Me.txtDddRecado.Text & ")"
                    Me.txtDddRecado.SelStart = Len(Me.txtDddRecado)
                End If
                SendKeys "{TAB}"
            Else
                MsgBox "O DDD deve conter dois dígitos!"
            End If
        Else
            SendKeys "{TAB}"
        End If
    End If
    
End Sub



Private Sub txtDigito_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Me.txtRg.MaxLength = 1

End Sub

Private Sub txtId_Change()

'Dim As Long
'codigo = Me.txtId.Value
'
'Me.txtNome = Application.WorksheetFunction.VLookup(codigo, Sheets("Planilha1").Range("A:BB"), 2, 0)
'


End Sub

Private Sub txtNumero_Change()

End Sub

Private Sub txtPlaca2_Change()

End Sub

Private Sub txtRenavam1_Change()

End Sub

Private Sub txtRg_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.txtRg.MaxLength = 8

End Sub

Private Sub txtTelefone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.txtTelefone.MaxLength = 9
    
End Sub

Private Sub txtTelefone2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.txtTelefone2.MaxLength = 9

End Sub

Private Sub txtTelefoneRecado_Change()


    Me.txtTelefoneRecado.MaxLength = 9
    

End Sub

Private Sub txtVencimentoCnh_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.txtVencimentoCnh.MaxLength = 10
    
    If Len(Me.txtVencimentoCnh) = 2 Then
    Me.txtVencimentoCnh.Text = Me.txtVencimentoCnh.Text & "/"
    Me.txtVencimentoCnh.SelStart = Len(Me.txtVencimentoCnh)
    End If
    
    If Len(Me.txtVencimentoCnh) = 5 Then
    Me.txtVencimentoCnh.Text = Me.txtVencimentoCnh.Text & "/"
    Me.txtVencimentoCnh.SelStart = Len(Me.txtVencimentoCnh)
    End If

End Sub

Private Sub DefinirFocoPrimeiraPagina()
    Me.MultiPage1.Value = 0 ' Define a primeira página como ativa
    Me.MultiPage1.SetFocus ' Define o foco em um controle da primeira página
End Sub

Private Sub VerificarCamposPreenchidos()

    ' impedir que salve caso o nome esteja vazio
'    If Me.txtNome = "" Then
'        MsgBox "Preencha o Nome!"
'        Exit Sub
'    End If

    Dim campos As Variant
    campos = Array(Me.txtNome, Me.txtNacionalidade, Me.txtCpf, Me.txtRg, _
                   Me.txtNumero, Me.txtCep, Me.txtComplemento, Me.txtDataNascimento, _
                   Me.cbbEstadoCivil, Me.txtProfissao, Me.txtTelefone, _
                   Me.txtEmail, Me.txtServico, Me.txtValor, _
                   Me.cbbPagamento, Me.txtCondicoesParce)

    
    Dim campo As Variant
    For Each campo In campos
        If campo.Value = "" Then
            MsgBox "Preencha todos os campos obrigatórios!"
            DefinirFocoPrimeiraPagina
            Exit Sub ' Sai do procedimento se houver campo vazio
        End If
    Next campo
End Sub


Private Sub UserForm_Activate()

    ' ajustando form para preencher toda a tela
    With UserForm1
        Width = Application.Width
        Height = Application.Height
        Left = Application.Left
        Top = Application.Top
     End With
     
     Me.Lista.RowSource = "Clientes"
     Me.Lista.ColumnCount = 54
     Me.Lista.ColumnHeads = True
    
     
    
End Sub

Private Sub UserForm_Initialize()

    ' desabilita as abas das páginas subsequentes, exceto pesquisar
    Dim i As Integer
    For i = 1 To Me.MultiPage1.Pages.Count - 1
        If Me.MultiPage1.Pages(i).name <> "pesquisar" Then
            Me.MultiPage1.Pages(i).Enabled = False
        End If
    Next i

    Dim wsId As Worksheet
    Set wsId = ThisWorkbook.Sheets("Gerar ID")

    Call LimparCampos
    
    Me.txtId.Value = wsId.Cells(2, 1).Value
    
    ' txtNome.SetFocus
    
    ' inserindo dados nos selects
    Me.cbbEstadoCivil.RowSource = "Dados!A2:A5"
    Me.cbbCategoria.RowSource = "Dados!C2:C10"
    Me.cbbPagamento.RowSource = "Dados!E2:E3"
    Me.cbbUfRg.RowSource = "Dados!G2:G29"
    
    ' ocultar campos variáveis
    lblRenavam2.Visible = False
    txtRenavam2.Visible = False
    lblPlaca2.Visible = False
    txtPlaca2.Visible = False
    
    lblRenavam3.Visible = False
    txtRenavam3.Visible = False
    lblPlaca3.Visible = False
    txtPlaca3.Visible = False
    
    lblPaCassacao2.Visible = False
    txtPaCassacao2.Visible = False
    lblPaCassacao3.Visible = False
    txtPaCassacao3.Visible = False
    
    lblPaSuspensao2.Visible = False
    txtPaSuspensao2.Visible = False
    lblPaSuspensao3.Visible = False
    txtPaSuspensao3.Visible = False

End Sub
