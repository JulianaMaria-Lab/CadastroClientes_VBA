Attribute VB_Name = "API_CEP"
Option Explicit
Function BuscarCEP(cep As String) As String

    Dim req As New MSXML2.ServerXMLHTTP60
    Dim endPoint As String
    
    endPoint = "https://viacep.com.br/ws/" & cep & "/json/"
    req.Open "GET", endPoint
    req.Send
    
    If req.Status = 200 Then
        Dim responseJson As Object
        Set responseJson = JsonConverter.ParseJson(req.ResponseText)
        
        If Not responseJson Is Nothing And responseJson("erro") = True Then
            BuscarCEP = "CEP N�O ENCONTRADO!"
        Else
            BuscarCEP = req.ResponseText
        End If
    Else
        BuscarCEP = "CEP N�O ENCONTRADO!"
    End If

End Function

Function Parse(resposta As String, indice As Integer)
    
    Dim matriz As Variant
    Dim subMatriz As Variant
    Dim resultado As String
    Dim i As Integer
    
    matriz = Split(resposta, ":")
    
    If indice >= 0 And indice < UBound(matriz) Then
        subMatriz = Split(matriz(indice), ",")
        resultado = subMatriz(0)
        resultado = Trim(resultado)
        resultado = Replace(resultado, Chr(34), "")
        Parse = resultado
    Else
        Parse = ""
    End If
     
End Function

Sub LimparCampos()

    Dim icontrol As MSForms.Control
    
    For Each icontrol In UserForm1.Controls
      If icontrol.Tag = "LB" Then
        icontrol.Caption = ""
      End If
        
    Next
    
End Sub

Function CampoVazio(Form As UserForm) As Boolean
    CampoVazio = False
    Dim Controle As Control
    Dim controlName As String
    Dim camposFaltando As String
    Dim primeiroCampoFaltando As Control
    Dim sequenciaCampos As Variant
    Dim i As Integer
    
    camposFaltando = ""
    Set primeiroCampoFaltando = Nothing
    
    ' sequ�ncia dos campos obrigat�rios
    sequenciaCampos = Array("txtNome", "txtCpf", "txtRg", "txtDigito", "cbbUfRg", "txtNacionalidade", "txtDataNascimento", "cbbEstadoCivil", "txtProfissao", "txtCep", "txtNumero", "txtDdd1", "txtTelefone", "txtEmail", "txtServico", "txtValor", "cbbPagamento")
    
    For i = LBound(sequenciaCampos) To UBound(sequenciaCampos)
        controlName = sequenciaCampos(i)
        For Each Controle In Form.Controls
            ' Verifica se o controle pertence � nova Multipage
            If Controle.Parent.name = "pesquisar" Then
                Controle.Enabled = True ' Habilita o controle na nova Multipage
                Exit For ' Sai do loop para evitar a valida��o nos outros controles
            End If
            
            If Controle.Tag = "campoObrigatorio" And Controle.name = controlName Then
                If TypeName(Controle) = "TextBox" Or TypeName(Controle) = "ComboBox" Then
                    If Controle.Value = "" Then
                        Controle.BackColor = RGB(255, 215, 215)
                        camposFaltando = camposFaltando & Controle.name & ", "
                        
                        If primeiroCampoFaltando Is Nothing Then
                            Set primeiroCampoFaltando = Controle
                        End If
                    Else
                        Controle.BackColor = VBA.vbWhite ' Redefine a cor de fundo para branco quando o campo n�o estiver vazio
                    End If
                End If
            End If
        Next Controle
    Next i
    
    If camposFaltando <> "" Then
        MsgBox "PREENCHA OS CAMPOS OBRIGAT�RIOS!", vbCritical, "CAMPOS OBRIGAT�RIOS"
        CampoVazio = True
        
        If Not primeiroCampoFaltando Is Nothing Then
            primeiroCampoFaltando.SetFocus ' Move o foco do cursor para o primeiro campo que est� faltando
        End If
    End If
End Function
