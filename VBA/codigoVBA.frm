VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProdutos 
   Caption         =   "Cadastro"
   ClientHeight    =   6330
   ClientLeft      =   -165
   ClientTop       =   -585
   ClientWidth     =   8730.001
   OleObjectBlob   =   "codigoVBA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub buttonCadastrar_Click()
    
    '1. criar as variaveis
    Dim codigo As Long, descricao As String, categoria As String
    Dim media As String, ano As Integer, classificacao As String
    Dim preco As Double, genero As String, dev As String, estoque As Integer
    Dim plataforma As String, nome As String, linha As Integer
    
    
    '2. inspecionar os preenchimentos e colocar vermelho
    If Not IsNumeric(textCodigo.Text) Then
        MsgBox "Favor preencher corretamente o campo código"
        Exit Sub
    End If
    
    If textNome.Text = "" Then
        MsgBox "Favor preencher o campo Descrição"
        textNome.BackColor = &HFF&
        Exit Sub
    End If
    
    If comboCategoria.Text = "" Then
        MsgBox "Favor preencher o campo Descrição"
        comboCategoria.BackColor = &HFF&
        Exit Sub
    End If
    
    If comboMidia.Text = "" Then
        MsgBox "Favor preencher o campo Descrição"
        comboMidia.BackColor = &HFF&
        Exit Sub
    End If
    
    If Not IsNumeric(textAno.Text) Then
        MsgBox "Favor preencher o campo Corretamente"
        textAno.BackColor = &HFF&
        Exit Sub
    End If
    
    If comboClassif.Text = "" Then
        MsgBox "Favor preencher o campo Descrição"
        comboClassif.BackColor = &HFF&
        Exit Sub
    End If
    
    
    If textDesc.Text = "" Then
        MsgBox "Favor preencher o campo Descrição"
        textDesc.BackColor = &HFF&
        Exit Sub
    End If
    
    If Not IsNumeric(textPreco.Text) Then
       MsgBox "Favor preencher o campo Corretamente"
       textPreco.BackColor = &HFF&
        Exit Sub
    End If
    
    If comboGen.Text = "" Then 'sda
        MsgBox "Favor preencher o campo Categoria"
        comboGen.BackColor = &HFF&
        Exit Sub
    End If
    
    If comboPlataform.Text = "" Then
        MsgBox "Favor preencher o campo Categoria"
        comboPlataform.BackColor = &HFF&
        Exit Sub
    End If
    
    'verificar se media for fisica e depois liberar estoque
    If comboMidia = "Fisica" Then
    
        If Not IsNumeric(TextEstoque.Text) Then
        
            MsgBox "Favor preencher o campo Corretamente"
            TextEstoque.BackColor = &HFF&
        
            
        End If
            
        If TextEstoque.Value < 1 Then
            
                MsgBox "Favor preencher o campo Corretamente"
                TextEstoque.BackColor = &HFF&
            
        Exit Sub
        End If

    End If
    
    
    
    
    
    If textDev.Text = "" Then
        MsgBox "Favor preencher o campo Categoria"
        textDev.BackColor = &HFF&
        Exit Sub
    End If
    
    
    
    'textPreco
    'comboGen
    'comboPlataform
         
     
     '3. passar os dados do formulario para as variaves
     codigo = textCodigo.Text
     nome = textNome.Text
     categoria = comboCategoria.Text
     descricao = textDesc.Text
     media = comboMidia.Text
     ano = textAno.Text
     classificacao = comboClassif.Text
     preco = textPreco.Text
     genero = comboGen.Text
     dev = textDev.Text
     
     If comboMidia = "Fisica" Then
        estoque = TextEstoque.Text
     End If
     
     plataforma = comboPlataform.Text
     
     
     '4. calcular o valor total
     'valorTotal = valor * qtdEstoque
     
     '5. pegar a linha da planilha de controle
     linha = planControle.Range("A2").Value
     'linha = linha + 1
     
     '6. passar os dados das variáveis para a planilha de produtos
     PlanProdutos.Cells(linha, 2).Value = codigo
     PlanProdutos.Cells(linha, 3).Value = nome
     PlanProdutos.Cells(linha, 4).Value = categoria
     PlanProdutos.Cells(linha, 5).Value = descricao
     PlanProdutos.Cells(linha, 6).Value = media
     PlanProdutos.Cells(linha, 7).Value = ano
     PlanProdutos.Cells(linha, 8).Value = classificacao
     PlanProdutos.Cells(linha, 9).Value = preco
     PlanProdutos.Cells(linha, 10).Value = genero
     PlanProdutos.Cells(linha, 11).Value = dev
     
     If estoque = 0 Then
        PlanProdutos.Cells(linha, 12).Value = "NULL"
     End If
     If estoque > 0 Then
        PlanProdutos.Cells(linha, 12).Value = estoque
     End If
     
     PlanProdutos.Cells(linha, 13).Value = plataforma
     
     '7. mudar a numeracao da linha
     linha = linha + 1
     planControle.Range("A2").Value = linha
     
     '8. Limpar os dados do formulario
     textCodigo.Value = ""
     textNome.Text = ""
     comboCategoria.Value = "" 'sdas
     textDesc.Text = ""
     comboMidia.Value = "" 'adfdf
     textAno.Text = ""
     comboClassif.Value = "" 'dfgd
     textPreco.Text = ""
     comboGen.Value = "" 'fsdgdg
     textDev.Text = ""
     TextEstoque.Text = ""
     comboPlataform.Value = "" 'dghd
     
     'Deixar todos compos brancos
     textCodigo.BackColor = &H8000000F

     textNome.BackColor = &H8000000F

     comboCategoria.BackColor = &H8000000F

     textDesc.BackColor = &H8000000F

     comboMidia.BackColor = &H8000000F

     textAno.BackColor = &H8000000F

     comboClassif.BackColor = &H8000000F

     textPreco.BackColor = &H8000000F

     comboGen.BackColor = &H8000000F

     textDev.BackColor = &H8000000F

     TextEstoque.BackColor = &H8000000F

     comboPlataform.BackColor = &H8000000F
     
     TextEstoque.BackColor = &HE0E0E0
     TextEstoque.Enabled = False
     comboGen.Enabled = False
     
     '9. colocar o foco no primeiro controle
     textCodigo.SetFocus
     
     MsgBox "Produto cadastrado com sucesso", vbInformation, "Sucesso"
        
    
End Sub

Private Sub buttonSair_Click()
    'fechar o formulario
    Unload Me
End Sub







Private Sub comboCategoria_Change()
    
    'Um Switch para alterar a lista genero de cordo com opcão campo categoria
    
    comboGen.Enabled = True
    
    Dim primeiroValor As String
    
    
    Select Case comboCategoria.Value
    
      Case "RPG" '
        
        comboGen.Clear
        comboGen.AddItem "RPG de Ação"
        comboGen.AddItem "MMORPG"
        comboGen.AddItem "Rouguelikes"
        primeiroValor = "RPG de Ação"
        
      Case "AçãoAventura"
        
        comboGen.Clear
        comboGen.AddItem "Horror e Sobrevivência"
        comboGen.AddItem "Metroidvania"
        comboGen.AddItem "FPS"
        primeiroValor = "Horror e Sobrevivência"
        
      Case "Simulação"
      
        comboGen.Clear
        comboGen.AddItem "Construção"
        comboGen.AddItem "Gestão"
        comboGen.AddItem "Vida"
        comboGen.AddItem "Veículos"
        primeiroValor = "Construção"
        
      Case "Esportes"
      
            
        comboGen.Clear
        comboGen.AddItem "Futebol"
        comboGen.AddItem "Basquete"
        comboGen.AddItem "Volei"
        comboGen.AddItem "Corrida"
        primeiroValor = "Futebol"
               
      Case "Estratégia"
      
            
        comboGen.Clear
        comboGen.AddItem "Puzzle"
        comboGen.AddItem "RTS"
        comboGen.AddItem "MOBA"
        primeiroValor = "Puzzle"
               
      Case Else
        Debug.Print "Not between 1 and 10"
    End Select
    
    comboGen.Value = primeiroValor


End Sub



Private Sub comboMidia_Change()

    'um IF liberando o campo Estoque caso midia for Fisica

If comboMidia.Value = "Fisica" Then

    TextEstoque.BackColor = &HFFFFFF
    TextEstoque.Enabled = True
    
    

End If

If comboMidia.Value = "Digital" Then

    TextEstoque.BackColor = &HE0E0E0
    TextEstoque.Enabled = False
    TextEstoque.Value = ""

    

End If


End Sub


Private Sub UserForm_Activate()

    'criando as lista dos combos do coampo cadastro
    comboCategoria.AddItem "RPG"
    comboCategoria.AddItem "AçãoAventura"
    comboCategoria.AddItem "Simulação"
    comboCategoria.AddItem "Esportes"
    comboCategoria.AddItem "Estratégia"

    
    
    comboClassif.AddItem "Livre"
    comboClassif.AddItem "10"
    comboClassif.AddItem "12"
    comboClassif.AddItem "14"
    comboClassif.AddItem "16"
    comboClassif.AddItem "18"
    
    comboMidia.AddItem "Digital"
    comboMidia.AddItem "Fisica"
    
    comboPlataform.AddItem "PC"
    comboPlataform.AddItem "Xbox Series X"
    comboPlataform.AddItem "Xbox One"
    comboPlataform.AddItem "playstation 5"
    comboPlataform.AddItem "playstation 4"
    comboPlataform.AddItem "Nintendo Switch"
    comboPlataform.AddItem "Nintendo Wii U"
    comboPlataform.AddItem "Nintendo 3DS"
    
    

    Dim idValorCategoria As Integer
    

    Debug.Print "valor " + comboCategoria.Value
    
    
End Sub
