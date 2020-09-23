VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procura de CDs"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
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
      Height          =   4695
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8415
      Begin VB.OptionButton optCategoria 
         Caption         =   "Categoria"
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
         Left            =   5520
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optPreco 
         Caption         =   "Preço"
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
         Left            =   7080
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Default         =   -1  'True
         Height          =   285
         Left            =   6600
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstResultado 
         Height          =   2895
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15400958
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Título"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Artista"
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Categoria"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Preço"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Quantidade"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.OptionButton optArtista 
         Caption         =   "Artista"
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
         Left            =   4200
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optTitulo 
         Caption         =   "Título"
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
         Left            =   3000
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optCodigo 
         Caption         =   "Código"
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
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtTexto 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Resultado:"
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
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Procurar por:"
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
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Critério:"
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
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmProcura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdPesquisar_Click()
  lstResultado.ListItems.Clear
  If tabcds.RecordCount = 0 Then Exit Sub
  If txtTexto = "" Then
    tabcds.MoveFirst
    Do While Not tabcds.EOF
      Set temp = lstResultado.ListItems.Add(, "", tabcds!codigo, 0, 0)
      'On Error Resume Next
      temp.SubItems(1) = tabcds!titulo
      temp.SubItems(2) = tabcds!artista
      temp.SubItems(3) = tabcds!categoria
      temp.SubItems(4) = tabcds!preco
      temp.SubItems(5) = tabcds!quantidade
      tabcds.MoveNext
    Loop
    Exit Sub
  End If
  If optCodigo.Value = True Then
    crit = "codigo=" & txtTexto
  ElseIf optTitulo.Value = True Then
    crit = "titulo='" & txtTexto & "'"
  ElseIf optArtista.Value = True Then
    crit = "artista='" & txtTexto & "'"
  ElseIf optCategoria.Value = True Then
    crit = "categoria='" & txtTexto & "'"
  ElseIf optPreco.Value = True Then
    crit = "preco='" & txtTexto & "'"
  End If
  tabcds.FindFirst crit
  Do While Not tabcds.NoMatch
    Set temp = lstResultado.ListItems.Add(, "", tabcds!codigo, 0, 0)
    'On Error Resume Next
    temp.SubItems(1) = tabcds!titulo
    temp.SubItems(2) = tabcds!artista
    temp.SubItems(3) = tabcds!categoria
    temp.SubItems(4) = tabcds!preco
    temp.SubItems(5) = tabcds!quantidade
    tabcds.FindNext crit
  Loop
End Sub

Private Sub optArtista_Click()
  txtTexto.SetFocus
End Sub

Private Sub optCategoria_Click()
  txtTexto.SetFocus
End Sub

Private Sub optCodigo_Click()
  If Not IsNumeric(txtTexto) Then
    txtTexto = ""
  End If
  txtTexto.SetFocus
End Sub

Private Sub optPreco_Click()
  txtTexto = ""
  txtTexto.SetFocus
End Sub

Private Sub optTitulo_Click()
  txtTexto.SetFocus
End Sub

Private Sub txtTexto_GotFocus()
  txtTexto.SelStart = 0
  txtTexto.SelLength = Len(txtTexto)
End Sub

Private Sub txtTexto_KeyPress(KeyAscii As Integer)
  If optCodigo.Value = True Then
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
  End If
End Sub
