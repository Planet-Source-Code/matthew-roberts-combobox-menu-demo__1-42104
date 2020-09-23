VERSION 5.00
Begin VB.Form frmCustomMenu 
   BorderStyle     =   0  'None
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   1890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbMenuItem 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblMenuItem 
      Caption         =   "x"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblMenuItem 
      Caption         =   "x"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   525
      Width           =   1575
   End
   Begin VB.Label lblMenuItem 
      Caption         =   "x"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   315
      Width           =   1575
   End
   Begin VB.Label lblMenuItem 
      Caption         =   "x"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   120
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmCustomMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_SelectedValue As Variant


Public Property Get SelectedValue()
    SelectedValue = m_SelectedValue
End Property

Private Property Let SelectedValue(Value)
    m_SelectedValue = Value
End Property

Private Sub cmbMenuItem_Click(Index As Integer)

    ReturnValue cmbMenuItem(Index).Text
End Sub

Private Sub Form_Load()
    
    lblMenuItem(0).Caption = "&Reports"
    lblMenuItem(0).Tag = "Reports"
    lblMenuItem(1).Caption = "&Export"
    lblMenuItem(1).Tag = "Export"
    lblMenuItem(2).Caption = "&DateFormat"
    lblMenuItem(2).Tag = "DateFormat"
        
    'lblMenuItem(0).BackColor =
   ' lblMenuItem(1).BackColor = Me.ForeColor
    'lblMenuItem(2).BackColor = Me.ForeColor
        
    
    With cmbMenuItem(0)
        .List(0) = "MM/DD/YYYY"
        .List(1) = "MM/DD/YY"
        .List(2) = "M/D/YYYY"
        .List(3) = "Mmm/DD/YYYY"
        .Text = "MM/DD/YYYY"
    End With
    
End Sub

Private Sub Form_LostFocus()
    Hide
End Sub

Private Sub lblMenuItem_Click(Index As Integer)
    ReturnValue lblMenuItem(Index).Tag
End Sub

Private Sub lblMenuItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    For X = 0 To lblMenuItem.Count - 1
        lblMenuItem(X).BackColor = Me.BackColor
        lblMenuItem(X).ForeColor = SystemColorConstants.vbWindowText
    Next X
    
    With lblMenuItem(Index)
        .BackColor = SystemColorConstants.vbHighlight
        .ForeColor = SystemColorConstants.vbHighlightText
    End With
    
    
End Sub


Private Sub ReturnValue(Value)

    SelectedValue = Value
    Hide

End Sub
