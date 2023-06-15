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
   Begin VB.ListBox listSnippet 
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
      Begin VB.Menu mEditor 
         Caption         =   "Editor"
         Begin VB.Menu mSnippet 
            Caption         =   "Snippet"
         End
         Begin VB.Menu mTexto 
            Caption         =   "Texto"
            Visible         =   0   'False
            Begin VB.Menu mEditar 
               Caption         =   "Editar"
            End
            Begin VB.Menu mSalvar 
               Caption         =   "Salvar"
            End
            Begin VB.Menu mAdicionar 
               Caption         =   "Adicionar"
            End
            Begin VB.Menu mRemover 
               Caption         =   "Remover"
            End
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

Private Sub Form_Load()
   ' Titulo do formulário
   Me.Caption = App.Title & "_v" & App.Major & "." & App.Minor & " by DALÇÓQUIO AUTOMAÇÃO"
   
   ' ToolTipText
   listSnippet.ToolTipText = "Duplo click para copiar."
   txtSnippet.ToolTipText = "Dublo click para menu."
   
   ' Mensagem de texto
   txtMensagem.Visible = False
   txtMensagem.Text = "Snippet copiado com sucesso..." & vbCrLf & _
                      "Use (Ctrl+V) no local desejado."
   
   ' Largura inicial do formulário
   Me.Width = 3600
   
   ' Carrega lista de snippets
   Call LoadSnippets
   
   ' Ordem Alfabética para lista de snippets
   Call OrdenarListBoxAlfabeticamente(listSnippet)
   
'Recupera os valores em config.ini
   Color = ReadIniValue(App.Path & "\Config.ini", "VARIAVEIS", "Color")
   
   ' Atualiza Color do Formulário
   If Color = "Black" Then Call mBlack_Click
   If Color = "White" Then Call mWhite_Click
   
End Sub

Private Sub mBlack_Click()
   ' Color Black
   mBlack.Checked = True
   mWhite.Checked = False
   Color = "Black"
   listSnippet.BackColor = vbBlack ' cor de fundo
   listSnippet.ForeColor = vbWhite  ' cor do texto
   txtSnippet.BackColor = vbBlack ' cor de fundo
   txtSnippet.ForeColor = vbWhite  ' cor do texto
   WriteIniValue App.Path & "\Config.ini", "VARIAVEIS", "Color", Color
   
 End Sub
 
Private Sub mWhite_Click()
   ' Color White
   mWhite.Checked = True
   mBlack.Checked = False
   Color = "White"
   listSnippet.BackColor = vbWhite ' cor de fundo
   listSnippet.ForeColor = vbBlack  ' cor do texto
   txtSnippet.BackColor = vbWhite ' cor de fundo
   txtSnippet.ForeColor = vbBlack  ' cor do texto
   WriteIniValue App.Path & "\Config.ini", "VARIAVEIS", "Color", Color

End Sub

Private Sub mSnippet_Click()
   If Me.Width = 3600 Then
      Me.Width = 12250 ' open
   Else
      Me.Width = 3600 ' close
      txtSnippet.Text = Empty
   End If

End Sub

Private Sub mHelp_Click()
    Dim filePath As String
    filePath = App.Path & "\help.html" ' Substitua pelo caminho do arquivo HTML desejado

    ' Abre o arquivo HTML no navegador padrão
    Shell "rundll32.exe url.dll,FileProtocolHandler " & filePath, vbNormalFocus
End Sub

'Private Sub txtSnippets_DblClick()
'   Dim projectPath As String
'   projectPath = App.Path
'
'   Shell "explorer.exe " & projectPath, vbNormalFocus
'
'End Sub

Private Sub txtSnippet_DblClick()
   PopupMenu mTexto
   
End Sub

Private Sub listSnippet_DblClick()
   ' Verifica se snippet selecionado
   If listSnippet.SelCount = 0 Then ' ou listSnippet.ListIndex >= 0
      MsgBox "Nenhum snippet selecionado para copiar", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
      
   Dim snippetName As String
   snippetName = listSnippet.List(listSnippet.ListIndex)
   
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

Private Sub mAdicionar_Click()
   Dim snippetName As String
   Dim snippetText As String
   
   ' Verifica se ah texto para snippet
   If txtSnippet.Text = Empty Then
      MsgBox "Digite o texto do snippet antes de adicionar.", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
   
   snippetName = InputBox("Digite o nome do snippet:", "DALÇÓQUIO AUTOMAÇÃO")
   ' verifica se o nome do snippet já existe
   If checkName(snippetName) = True Then
      MsgBox "Nome para snippet já existente !!!", vbExclamation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
   
   ' Verifica se tem nome para o snippet
   If snippetName <> Empty Then
     snippetText = txtSnippet.Text
         ' Adiciona o nome do snippet à lista
         listSnippet.AddItem snippetName
         
         ' Salva o texto do snippet em um arquivo
         Call SaveSnippet(snippetName, snippetText)
         
         ' Limpa o TextBox
         txtSnippet.Text = Empty
   Else
      MsgBox "Nome para snippet em branco ou cancelado.", vbExclamation, "DALÇÓQUIO AUTOMAÇÃO"
      End If
 
End Sub

Private Sub mRemover_Click()
    ' Verifica se snippet selecionado
    If listSnippet.SelCount = 0 Then ' ou If listSnippet.ListIndex >= 0 Then
        MsgBox "Nenhum snippet selecionado para remover", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        Exit Sub
    End If
    
    Dim snippetName As String
    snippetName = listSnippet.List(listSnippet.ListIndex)

    ' Confirmação do usuário
    Dim response As VbMsgBoxResult
    response = MsgBox("Tem certeza de que deseja remover o snippet selecionado?", vbYesNo + vbQuestion, "DALÇÓQUIO AUTOMAÇÃO")

    If response = vbYes Then
        ' Remove o snippet da lista
        listSnippet.RemoveItem listSnippet.ListIndex

        ' Exclui o arquivo de texto do snippet
        DeleteSnippetFile snippetName

        ' Limpa o TextBox
        txtSnippet.Text = Empty
    End If
   
End Sub

Private Sub mEditar_Click()
   ' Verifica se snippet selecionado
   If listSnippet.SelCount = 0 Then ' ou listSnippet.ListIndex >= 0
      MsgBox "Nenhum snippet selecionado para editar", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
      
   ' Obtém o nome do snippet
   Dim snippetName As String
   snippetName = listSnippet.List(listSnippet.ListIndex)
   
   ' Obtém o texto do snippet do arquivo
   Dim snippetText As String
   snippetText = ReadSnippet(snippetName)
   
   txtSnippet.Text = snippetText

End Sub

Private Sub mSalvar_Click()
   Dim snippetText As String
   Dim snippetName As String
   
   ' Verifica se snippet selecionado
   If listSnippet.SelCount = 0 Then ' ou listSnippet.ListIndex >= 0
      MsgBox "Nenhum snippet selecionado para salvar", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
   
   ' Verifica se ah texto para snippet
   If txtSnippet.Text = Empty Then
      MsgBox "Digite o texto do snippet antes de salvar.", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
   
   ' Confirmação do usuário
    Dim response As VbMsgBoxResult
    response = MsgBox("Tem certeza de que deseja salvar o snippet selecionado?", vbYesNo + vbQuestion, "DALÇÓQUIO AUTOMAÇÃO")
    If response = vbNo Then Exit Sub
   
   snippetName = listSnippet.List(listSnippet.ListIndex)
   
   ' Exclui o arquivo de texto do snippet para criar um atualizado
   DeleteSnippetFile snippetName
   
   'Salva o texto do snippet em um arquivo
   snippetText = txtSnippet.Text
   Call SaveSnippet(snippetName, snippetText)
           
   ' Limpa o TextBox
   txtSnippet.Text = Empty

End Sub

Private Sub LoadSnippets()
   Dim fileName As String
   fileName = App.Path & "\snippets.txt"
   
   If Dir(fileName) <> "" Then
       Dim snippetName As String
       Open fileName For Input As #1
       Do Until EOF(1)
           Line Input #1, snippetName
           listSnippet.AddItem snippetName
       Loop
       Close #1
   End If
End Sub

Private Sub SaveSnippet(ByVal snippetName As String, ByVal snippetText As String)
   Dim fileName As String
   fileName = App.Path & "\" & snippetName & ".txt"
   
   Open fileName For Output As #1
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
   Dim fileName As String
   fileName = App.Path & "\" & snippetName & ".txt"
   
   If Dir(fileName) <> Empty Then
       Open fileName For Input As #1
       ReadSnippet = Input$(LOF(1), 1)
       Close #1
   Else
       ReadSnippet = Empty
   End If
   
End Function

Private Sub DeleteSnippetFile(ByVal snippetName As String)
   Dim fileName As String
   fileName = App.Path & "\" & snippetName & ".txt"
   
   If Dir(fileName) <> Empty Then
       Kill fileName
   End If
   
   ' Remove o nome do snippet do arquivo de snippets
   Dim snippetsFileName As String
   snippetsFileName = App.Path & "\snippets.txt"
   
   If Dir(snippetsFileName) <> "" Then
       Dim tempFileName As String
       tempFileName = App.Path & "\temp.txt"
       
       Open snippetsFileName For Input As #1
       Open tempFileName For Output As #2
       
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
       Name tempFileName As snippetsFileName
   End If
   
End Sub

Function checkName(itemName As String) As Boolean
   
   Dim itemExists As Boolean
   itemExists = False
   
   Dim i As Integer
   For i = 0 To listSnippet.ListCount - 1
       If listSnippet.List(i) = itemName Then
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



Private Sub Timer1_Timer()
   ' Fecha texto de mensagem
   txtMensagem.Visible = False
   Timer1.Enabled = False

End Sub
