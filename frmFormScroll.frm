VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scroll Form"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2115
      Left            =   3300
      Max             =   10
      TabIndex        =   1
      Top             =   30
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   7155
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3195
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmFormScroll.frx":0000
         Top             =   2460
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "End of Transmition"
         Height          =   195
         Left            =   900
         TabIndex        =   6
         Top             =   6600
         Width           =   1320
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   3660
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Well, rahter than have tabs upon tabs, or frames upon frame, why not just scroll up and down as if it were a text box?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   960
         TabIndex        =   5
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Or this Circle"
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   3360
         Width           =   915
      End
      Begin VB.Shape Shape1 
         Height          =   1035
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2940
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "What would you do if you had a really small window, but had a whole bunch of crap you needed to put on it?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public frmHeight As Integer
Public frmtop As Integer
Public NeededTop As Integer

'To Increase / Decrease Scroll speed, edit VScroll1.Max
'   The Higher the MAX the finer the scrolling.

Private Sub Form_Load()
    frmHeight = Frame1.Height
    frmtop = Frame1.Top
    NeededTop = Form1.Height - (Frame1.Height) - 500
    VScroll1.Value = VScroll1.Min
    Label1 = "What would you do if you had a really small window, but had a whole bunch of crap you needed to put on it?"
End Sub

Private Sub VScroll1_Change()
    Frame1.Top = frmtop - (Abs(NeededTop / VScroll1.Max) * (VScroll1.Value))
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
