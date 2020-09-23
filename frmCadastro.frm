VERSION 5.00
Begin VB.Form frmCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de CDs"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   2220
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cboCategoria 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "frmCadastro.frx":0442
         Left            =   1440
         List            =   "frmCadastro.frx":0464
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtPreco 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtQuantidade 
         Height          =   285
         Left            =   4080
         TabIndex        =   5
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtArtista 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtTitulo 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Categoria:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Preço:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Quantidade:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Artista:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Título:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcluir_Click()
  tabcds.FindFirst "codigo=" & txtCodigo
  If tabcds.NoMatch Then
    MsgBox "Registro inexistente", vbOKOnly, "Registro inexistente"
  Else
    If MsgBox("Tem certeza que deseja excluir o registro?", vbYesNo, "Excluir registro") = vbYes Then
      tabcds.Delete
      ApagaTudo
    End If
  End If
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdNovo_Click()
  If txtCodigo <> "" Then
    tabcds.FindFirst "codigo=" & txtCodigo
    If tabcds.NoMatch Then
      If MsgBox("O registro atual ainda não foi salvo, deseja salvar?", vbYesNo, "Novo registro") = vbYes Then
        cmdSalvar_Click
      End If
    End If
  End If
  If tabcds.RecordCount = 0 Then Exit Sub
  tabcds.MoveLast
  temp = tabcds!codigo + 1
  ApagaTudo
  txtCodigo = temp
  txtTitulo.SetFocus
End Sub

Private Sub cmdSalvar_Click()
  tabcds.FindFirst "codigo=" & txtCodigo
  If tabcds.NoMatch Then
    tabcds.AddNew
  Else
    If MsgBox("Esse registro já existe, deseja sobrescreve-lo?", vbYesNo, "Registro existente") = vbYes Then
      tabcds.Edit
    Else
      Exit Sub
    End If
  End If
  tabcds!codigo = txtCodigo
  tabcds!titulo = txtTitulo
  tabcds!artista = txtArtista
  tabcds!categoria = cboCategoria
  tabcds!preco = txtPreco
  tabcds!quantidade = txtQuantidade
  tabcds.Update
End Sub

Private Sub Command1_Click()
  On Error GoTo PrinterError
  
  Printer.Line (1500, 500)-Step(5500, 2400), , B
    
  Printer.Font = "Tahoma"
  Printer.FontSize = 12
  AjustaXY 1700, 600
  Printer.Print "Código: " & txtCodigo
  AjustaXY 1700, 900
  Printer.Print "Título: " & txtTitulo
  AjustaXY 1700, 1200
  Printer.Print "Artista: " & txtArtista
  AjustaXY 1700, 1500
  Printer.Print "Categoria: " & cboCategoria
  AjustaXY 1700, 1800
  Printer.Print "Preço: " & txtPreco
  AjustaXY 1700, 2100
  Printer.Print "Quantidade: " & txtQuantidade
  AjustaXY 1700, 2500
  Printer.FontSize = 9
  Printer.Print "Dados imprimidos dia " & Date
  
  Printer.EndDoc
  
PrinterError:
End Sub

Private Sub Form_Load()
  Top = (frmPrincipal.ScaleHeight / 2) - (ScaleHeight / 2) - 200
  Left = (frmPrincipal.ScaleWidth / 2) - (ScaleWidth / 2)
End Sub

Private Sub txtArtista_GotFocus()
  txtArtista.SelStart = 0
  txtArtista.SelLength = Len(txtArtista)
End Sub

Private Sub txtCodigo_GotFocus()
  txtCodigo.SelStart = 0
  txtCodigo.SelLength = Len(txtCodigo)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 8 Then Exit Sub
  If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtCodigo_LostFocus()
  On Error Resume Next
  tabcds.FindFirst "codigo=" & txtCodigo
  If tabcds.NoMatch Then Exit Sub
  txtTitulo = tabcds!titulo
  txtArtista = tabcds!artista
  cboCategoria = tabcds!categoria
  txtPreco = tabcds!preco
  txtQuantidade = tabcds!quantidade
End Sub

Private Sub txtPreco_Change()
  txtPreco.SelStart = 0
  txtPreco.SelLength = Len(txtPreco)
End Sub

Private Sub txtQuantidade_GotFocus()
  txtQuantidade.SelStart = 0
  txtQuantidade.SelLength = Len(txtQuantidade)
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Sub ApagaTudo()
  txtCodigo = ""
  txtTitulo = ""
  txtArtista = ""
  cboCategoria = ""
  txtPreco = ""
  txtQuantidade = ""
End Sub

Private Sub txtTitulo_GotFocus()
  txtTitulo.SelStart = 0
  txtTitulo.SelLength = Len(txtTitulo)
End Sub

Sub AjustaXY(X As Integer, Y As Integer)
  Printer.CurrentX = X
  Printer.CurrentY = Y
End Sub
