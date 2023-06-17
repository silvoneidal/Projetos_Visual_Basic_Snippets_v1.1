VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   11130
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11130
   ScaleWidth      =   12150
   Begin VB.TextBox txtMensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   10440
      Width           =   3255
   End
   Begin VB.ListBox listSnippets 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   11085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   11400
      Top             =   10080
   End
   Begin VB.TextBox txtSnippet 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   11100
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   0
      Width           =   8655
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Begin VB.Menu mSnippet 
         Caption         =   "Snippet"
         Begin VB.Menu mAbrir 
            Caption         =   "Abrir"
         End
         Begin VB.Menu mSalvar 
            Caption         =   "Salvar"
         End
         Begin VB.Menu mExcluir 
            Caption         =   "Excluir"
         End
         Begin VB.Menu mRenomear 
            Caption         =   "Renomear"
         End
      End
      Begin VB.Menu mColor 
         Caption         =   "Color"
         Begin VB.Menu mBlack 
            Caption         =   "Black"
            Checked         =   -1  'True
         End
         Begin VB.Menu mWhite 
            Caption         =   "White"
         End
      End
      Begin VB.Menu mHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Color As String
Dim filePathSnippets As String
Dim filePathHelp As String

Private Sub Form_Load()
   ' Titulo do formulário
   Me.Caption = App.Title & "_v" & App.Major & "." & App.Minor & " by DALÇÓQUIO AUTOMAÇÃO"
   
   ' Local de arquivos
   filePathSnippets = App.Path & "\snippets.txt"
   filePathHelp = App.Path & "\help.html"
   
   ' ToolTipText
   listSnippets.ToolTipText = "Duplo click para copiar."
   
   ' Mensagem de texto
   txtMensagem.Visible = False
   txtMensagem.Text = "Snippet copiado com sucesso..." & vbCrLf & _
                      "Use (Ctrl+V) no local desejado."
   
   ' Largura inicial do formulário
   Me.Width = 3600
   
   ' Carrega lista de snippets
   Call LoadSnippets
   
'Recupera os valores em config.ini
   Color = ReadIniValue(App.Path & "\Config.ini", "VARIAVEIS", "Color")
   
   ' Atualiza Color do Formulário
   If Color = "Black" Then Call mBlack_Click
   If Color = "White" Then Call mWhite_Click
   
End Sub

Private Sub Timer1_Timer()
   ' Fecha texto de mensagem
   txtMensagem.Visible = False
   Timer1.Enabled = False

End Sub

Private Sub mAbrir_Click()
   If Me.Width = 3600 Then
      Me.Width = 12250 ' open
      mAbrir.Caption = "Fechar"
   Else
      Me.Width = 3600 ' close
      mAbrir.Caption = "Abrir"
   End If

End Sub

Private Sub mSalvar_Click()
   Dim snippetText As String
   Dim snippetName As String
   
   snippetName = listSnippets.List(listSnippets.ListIndex)
   
   ' Verifica se ah texto para snippet
   If txtSnippet.Text = Empty Then
      MsgBox "Digite um texto para o snippet antes de salvar.", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
      
   ' Verifica se ha snippet selecionado
   If listSnippets.SelCount > 0 Then ' ou listSnippets.ListIndex >= 0
      ' Confirmação do usuário
      Dim response As VbMsgBoxResult
      response = MsgBox("Deseja salvar no snippet: " & snippetName & " ?", vbYesNo + vbQuestion, "DALÇÓQUIO AUTOMAÇÃO")
      If response = vbYes Then
         GoTo SNIPPET_SELECT
      Else
         GoTo SNIPPET_NEW
      End If
   Else
      GoTo SNIPPET_NEW
   End If
   
SNIPPET_SELECT:
   ' Exclui o arquivo.txt do snippet
   DeleteSnippetFile snippetName
   
   'Salva o texto do snippet em um arquivo.txt
   snippetText = txtSnippet.Text
   Call SaveSnippet(snippetName, snippetText)
           
   ' Carrega lista de snippets
   Call LoadSnippets
   
   ' Confirmação de que o snippet foi salvo
   MsgBox "Snippet: " & snippetName & " salvo com sucesso...", , "DALÇÓQUIO AUTOMAÇÃO"
   Exit Sub

SNIPPET_NEW:
   snippetName = InputBox("Digite um nome para o snippet:", "DALÇÓQUIO AUTOMAÇÃO")
   ' verifica se o nome do snippet já existe
   If checkName(snippetName) = True Then
      MsgBox "Nome para snippet já existente !!!", vbExclamation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
   
   ' Verifica se tem nome para o snippet
   If snippetName <> Empty Then
      snippetText = txtSnippet.Text
      
      ' Adiciona o nome do snippet ao ListBox
      listSnippets.AddItem snippetName

      ' Salva o texto do snippet em um arquivo
      Call SaveSnippet(snippetName, snippetText)
      
      ' Carrega lista de snippets
      Call LoadSnippets

      ' Confirmação de que o snippet foi excluido
      MsgBox "Snippet: " & snippetName & " salvo com sucesso...", , "DALÇÓQUIO AUTOMAÇÃO"
   Else
      MsgBox "Nome para snippet em branco ou cancelado.", vbExclamation, "DALÇÓQUIO AUTOMAÇÃO"
   End If
 
End Sub

Private Sub mExcluir_Click()
    ' Verifica se snippet selecionado
    If listSnippets.SelCount = 0 Then ' ou If listSnippets.ListIndex >= 0 Then
        MsgBox "Nenhum snippet selecionado para excluir", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        Exit Sub
    End If
    
    Dim snippetName As String
    snippetName = listSnippets.List(listSnippets.ListIndex)

    ' Confirmação do usuário
    Dim response As VbMsgBoxResult
    response = MsgBox("Tem certeza de que deseja excluir o snippet: " & snippetName & " ?", vbYesNo + vbQuestion, "DALÇÓQUIO AUTOMAÇÃO")

    If response = vbYes Then
        ' Remove o snippet da lista
        listSnippets.RemoveItem listSnippets.ListIndex

        ' Exclui o arquivo de texto do snippet
        Call DeleteSnippetFile(snippetName)

        ' Limpa o TextBox
        txtSnippet.Text = Empty
    End If
   
End Sub

Private Sub mRenomear_Click()
   Dim snippetTemp As String
   Dim snippetName As String
   Dim snippetText As String
   
   ' Verifica se ah snippet selecionado
    If listSnippets.SelCount = 0 Then ' ou If listSnippets.ListIndex >= 0 Then
        MsgBox "Nenhum snippet selecionado para renomear", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        Exit Sub
    End If
    
    ' Verifica nome do snippet selecionado
    snippetName = listSnippets.List(listSnippets.ListIndex)
    ' Guarda temporáriamente o nome atual do snippet
    snippetTemp = snippetName
    
    ' Mensagem para o usuário
    snippetName = InputBox("Digite um novo nome para o snippet:", "DALÇÓQUIO AUTOMAÇÃO", snippetName)
   ' verifica se o nome do snippet já existe
   If checkName(snippetName) = True Then
      MsgBox "Nome para snippet já existente !!!", vbExclamation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
   
   ' Exclui o arquivo.txt do snippet
   Call DeleteSnippetFile(snippetTemp)
   
   'Salva novamente o texto do snippet em um arquivo.txt
   snippetText = txtSnippet.Text
   Call SaveSnippet(snippetName, snippetText)
           
   ' Carrega lista de snippets
   Call LoadSnippets
   
   ' Confirmação de que o snippet foi excluido
    MsgBox "Snippet: " & snippetTemp & " para " & snippetName & " renomeado com sucesso...", , "DALÇÓQUIO AUTOMAÇÃO"

End Sub

Private Sub mBlack_Click()
   ' Color Black
   mBlack.Checked = True
   mWhite.Checked = False
   Color = "Black"
   listSnippets.BackColor = vbBlack ' cor de fundo
   listSnippets.ForeColor = vbWhite  ' cor do texto
   txtSnippet.BackColor = vbBlack ' cor de fundo
   txtSnippet.ForeColor = vbWhite  ' cor do texto
   WriteIniValue App.Path & "\Config.ini", "VARIAVEIS", "Color", Color
   
 End Sub
 
Private Sub mWhite_Click()
   ' Color White
   mWhite.Checked = True
   mBlack.Checked = False
   Color = "White"
   listSnippets.BackColor = vbWhite ' cor de fundo
   listSnippets.ForeColor = vbBlack  ' cor do texto
   txtSnippet.BackColor = vbWhite ' cor de fundo
   txtSnippet.ForeColor = vbBlack  ' cor do texto
   WriteIniValue App.Path & "\Config.ini", "VARIAVEIS", "Color", Color

End Sub

Private Sub mHelp_Click()
    Dim filePath As String
    filePath = App.Path & "\help.html" ' Substitua pelo caminho do arquivo HTML desejado

    ' Abre o arquivo HTML no navegador padrão
    Shell "rundll32.exe url.dll,FileProtocolHandler " & filePath, vbNormalFocus
End Sub

Private Sub listSnippets_Click()
   ' Obtém o nome do snippet
   Dim snippetName As String
   snippetName = listSnippets.List(listSnippets.ListIndex)
   
   ' Obtém o texto do snippet do arquivo
   Dim snippetText As String
   snippetText = ReadSnippet(snippetName)
   
   txtSnippet.Text = snippetText
   
End Sub

Private Sub listSnippets_DblClick()
   ' Verifica se snippet selecionado
   If listSnippets.SelCount = 0 Then ' ou listSnippets.ListIndex >= 0
      MsgBox "Nenhum snippet selecionado para copiar", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
      
   Dim snippetName As String
   snippetName = listSnippets.List(listSnippets.ListIndex)
   
   ' Obtém o texto do snippet do arquivo
   Dim snippetText As String
   snippetText = ReadSnippet(snippetName)
   
   ' Copia o texto do snippet para a área de transferência
   Clipboard.Clear
   Clipboard.SetText snippetText
   
   Timer1.Enabled = True
   txtMensagem.Visible = True
   'MsgBox "O snippet foi copiado para a área de transferência (Ctrl+V para colar).", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
   
End Sub


Private Sub LoadSnippets()
   ' Limpa o ListBox
   listSnippets.Clear
   
   If Dir(filePathSnippets) <> "" Then
       Dim snippetName As String
       Open filePathSnippets For Input As #1
       Do Until EOF(1)
           Line Input #1, snippetName
           listSnippets.AddItem snippetName
       Loop
       Close #1
       ' Ordem Alfabética para lista de snippets
       Call OrdenarListBoxAlfabeticamente(listSnippets)
   End If
End Sub

Private Sub SaveSnippet(ByVal snippetName As String, ByVal snippetText As String)
   Dim filePathSnippet As String
   filePathSnippet = App.Path & "\" & snippetName & ".txt"
   
   Open filePathSnippet For Output As #1
   Print #1, snippetText
   Close #1
   
   ' Salva o nome do snippet no arquivo de snippets
   Dim snippetsFileName As String
   snippetsFileName = App.Path & "\snippets.txt"
   
   Open snippetsFileName For Append As #2
   Print #2, snippetName
   Close #2
   
End Sub

Private Function ReadSnippet(ByVal snippetName As String) As String
   Dim filePathSnippet As String
   filePathSnippet = App.Path & "\" & snippetName & ".txt"
   
   If Dir(filePathSnippet) <> Empty Then
       Open filePathSnippet For Input As #1
       ReadSnippet = Input$(LOF(1), 1)
       Close #1
   Else
       ReadSnippet = Empty
   End If
   
End Function

Private Sub DeleteSnippetFile(ByVal snippetName As String)
   Dim filePathSnippet As String
   filePathSnippet = App.Path & "\" & snippetName & ".txt"
   
   If Dir(filePathSnippet) <> Empty Then
       Kill filePathSnippet
   End If
   
   ' Remove o nome do snippet do arquivo de snippets
   Dim snippetsFileName As String
   snippetsFileName = App.Path & "\snippets.txt"
   Dim filePathTemp As String
   filePathTemp = App.Path & "\temp.txt"
   
   If Dir(snippetsFileName) <> "" Then
       Open snippetsFileName For Input As #1
       Open filePathTemp For Output As #2
       
       Do Until EOF(1)
           Dim line As String
           Line Input #1, line
           
           If Trim(line) <> snippetName Then
               Print #2, line
           End If
       Loop
       
       Close #1
       Close #2
       
       Kill snippetsFileName
       Name filePathTemp As snippetsFileName
   End If
   
End Sub

Function checkName(itemName As String) As Boolean
   
   Dim itemExists As Boolean
   itemExists = False
   
   Dim i As Integer
   For i = 0 To listSnippets.ListCount - 1
       If listSnippets.List(i) = itemName Then
           ' O item com o mesmo nome foi encontrado
           itemExists = True
           Exit For
       End If
   Next i
   
   If itemExists Then
       checkName = True ' já existe snippet com este nome
   Else
       checkName = False ' não existe snippet com este nome
   End If
   
End Function

Private Sub OrdenarListBoxAlfabeticamente(lstBox As ListBox)
    Dim arrItens() As String
    Dim i As Integer

    ' Armazena os itens do ListBox em um array
    ReDim arrItens(lstBox.ListCount - 1)
    For i = 0 To lstBox.ListCount - 1
        arrItens(i) = lstBox.List(i)
    Next i

    ' Ordena o array em ordem alfabética
    Call QuickSort(arrItens, 0, UBound(arrItens))

    ' Limpa o ListBox
    lstBox.Clear

    ' Adiciona os itens ordenados de volta ao ListBox
    For i = 0 To UBound(arrItens)
        lstBox.AddItem arrItens(i)
    Next i
End Sub

Private Sub QuickSort(arr() As String, left As Integer, right As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim pivot As String
    Dim Temp As String

    i = left
    j = right
    pivot = arr((left + right) \ 2)

    While i <= j
        While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Wend
        While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Wend
        If i <= j Then
            Temp = arr(i)
            arr(i) = arr(j)
            arr(j) = Temp
            i = i + 1
            j = j - 1
        End If
    Wend

    If left < j Then
        QuickSort arr, left, j
    End If
    If i < right Then
        QuickSort arr, i, right
    End If
End Sub



