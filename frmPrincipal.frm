VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmPrincipal 
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Vendas de CDs"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7125
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6090
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
            Text            =   "Sistema de Vendas de CDs"
            TextSave        =   "Sistema de Vendas de CDs"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "31/07/01"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "23:57"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBotoes 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7125
      TabIndex        =   0
      Top             =   0
      Width           =   7125
      Begin VB.Line lnbt 
         BorderColor     =   &H00808080&
         Index           =   2
         Visible         =   0   'False
         X1              =   1340
         X2              =   1340
         Y1              =   100
         Y2              =   620
      End
      Begin VB.Line lnbt 
         BorderColor     =   &H00808080&
         Index           =   3
         Visible         =   0   'False
         X1              =   820
         X2              =   1340
         Y1              =   620
         Y2              =   620
      End
      Begin VB.Line lnbt 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         Visible         =   0   'False
         X1              =   820
         X2              =   1340
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Line lnbt 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         Visible         =   0   'False
         X1              =   820
         X2              =   820
         Y1              =   100
         Y2              =   620
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   840
         Y1              =   10
         Y2              =   10
      End
      Begin VB.Image imgBotoes 
         Height          =   480
         Index           =   1
         Left            =   840
         Picture         =   "frmPrincipal.frx":0442
         ToolTipText     =   "Procura de CDs"
         Top             =   120
         Width           =   480
      End
      Begin VB.Image imgBotoes 
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "frmPrincipal.frx":0884
         ToolTipText     =   "Cadastro de CDs"
         Top             =   120
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   1200
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Principal"
      Begin VB.Menu mnuCadastro 
         Caption         =   "&Cadastro de CDs"
      End
      Begin VB.Menu mnuProcura 
         Caption         =   "&Procura de CDs"
      End
      Begin VB.Menu sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuAjuda2 
         Caption         =   "&Ajuda"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "&Sobre..."
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgBotoes_Click(Index As Integer)
  Select Case Index
    Case 0 ' Cadastro de CDs
      frmCadastro.Show
    Case 1 ' Procura de CDs
      frmProcura.Show
  End Select
End Sub

Private Sub imgBotoes_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Index
    Case 0 ' Cadastro de CDs
      AjustaXY 0, 100, 100, 100, 620
      AjustaXY 1, 100, 620, 100, 100
      AjustaXY 2, 620, 620, 100, 620
      AjustaXY 3, 120, 620, 620, 620
      LnBtVisible True
    Case 1 ' Procura de CDs
      AjustaXY 0, 820, 820, 120, 620
      AjustaXY 1, 820, 1340, 100, 100
      AjustaXY 2, 1340, 1340, 100, 620
      AjustaXY 3, 820, 1340, 620, 620
      LnBtVisible True
  End Select
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  LnBtVisible False
End Sub

Private Sub MDIForm_Resize()
  Line1.X1 = 0: Line1.X2 = Me.Width: Line1.Y1 = 0: Line1.Y2 = 0
  Line2.X1 = 0: Line2.X2 = Me.Width: Line2.Y1 = 10: Line2.Y2 = 10
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  tabcds.Close
  banco.Close
End Sub

Private Sub mnuCadastro_Click()
  frmCadastro.Show
End Sub

Private Sub mnuProcura_Click()
  frmProcura.Show
End Sub

Private Sub mnuSair_Click()
  Unload Me
End Sub

Private Sub mnuSobre_Click()
  frmSobre.Show 1, Me
End Sub

Sub AjustaXY(Index As Integer, X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer)
  With lnbt(Index)
    .X1 = X1
    .X2 = X2
    .Y1 = Y1
    .Y2 = Y2
  End With
End Sub

Sub LnBtVisible(Visivel As Boolean)
  lnbt(0).Visible = Visivel
  lnbt(1).Visible = Visivel
  lnbt(2).Visible = Visivel
  lnbt(3).Visible = Visivel
End Sub

Private Sub picBotoes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  LnBtVisible False
End Sub
