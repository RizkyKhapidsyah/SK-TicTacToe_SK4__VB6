VERSION 5.00
Begin VB.Form frmcttt 
   BackColor       =   &H00000000&
   Caption         =   "PLAY TIC TAC TOE WITH THE COMPUTER"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   DrawWidth       =   9666
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "STATUS"
      Height          =   975
      Left            =   120
      TabIndex        =   31
      Top             =   4440
      Width           =   8055
      Begin VB.Label LBL1 
         BackColor       =   &H00C0E0FF&
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "TIC TAC TOE: PLAY AGAINST THE COMPUTER"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF0000&
         Caption         =   "OPTIONS"
         Height          =   3735
         Left            =   6240
         TabIndex        =   21
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton Command2 
            BackColor       =   &H0000FFFF&
            Caption         =   "STATUS"
            Enabled         =   0   'False
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H0000FFFF&
            Caption         =   "EXIT"
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FF80FF&
            Caption         =   "RESTART"
            Enabled         =   0   'False
            Height          =   1455
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1455
            Begin VB.CommandButton Command12 
               BackColor       =   &H0000FFFF&
               Caption         =   "SAME PLAYER"
               Enabled         =   0   'False
               Height          =   495
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   840
               Width           =   1215
            End
            Begin VB.CommandButton Command10 
               BackColor       =   &H0000FFFF&
               Caption         =   "DIFFERENT PLAYER"
               Enabled         =   0   'False
               Height          =   495
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H000000FF&
            Caption         =   "START"
            Enabled         =   0   'False
            Height          =   735
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1455
            Begin VB.CommandButton Command11 
               BackColor       =   &H0000FFFF&
               Caption         =   "START"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   240
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF0000&
         Caption         =   "PLAYERS INFORMATION"
         Height          =   3735
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton Command13 
            BackColor       =   &H000000FF&
            Caption         =   "ALL INFORMATION GIVEN"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   3240
            Width           =   2175
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FF80FF&
            Caption         =   "LEVEL"
            Height          =   1935
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   2175
            Begin VB.OptionButton Option4 
               BackColor       =   &H00FFFF00&
               Caption         =   "KNOW THE GAME"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   1080
               Width           =   1815
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00FFFF00&
               Caption         =   "CALL ME CHAMPION"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   1440
               Width           =   1935
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00FFFF00&
               Caption         =   "HAVE AN IDEA OF IT"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   720
               Width           =   1935
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFF00&
               Caption         =   "NEW TO THE GAME"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Width           =   1935
            End
         End
         Begin VB.TextBox TXTAGE 
            BackColor       =   &H0000FF00&
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   600
            TabIndex        =   13
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox TXTNAME 
            BackColor       =   &H0000FF00&
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   600
            TabIndex        =   12
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "AGE"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "NAME"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF0000&
         Caption         =   "TIC TAC TOE"
         Enabled         =   0   'False
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3495
         Begin VB.PictureBox Picture9 
            Height          =   975
            Left            =   120
            Picture         =   "FRMCTTT.frx":0000
            ScaleHeight     =   915
            ScaleWidth      =   1035
            TabIndex        =   41
            Top             =   240
            Width           =   1095
         End
         Begin VB.PictureBox Picture8 
            Height          =   975
            Left            =   2280
            Picture         =   "FRMCTTT.frx":385E
            ScaleHeight     =   915
            ScaleWidth      =   1035
            TabIndex        =   40
            Top             =   1200
            Width           =   1095
         End
         Begin VB.PictureBox Picture7 
            Height          =   975
            Left            =   2280
            Picture         =   "FRMCTTT.frx":70BC
            ScaleHeight     =   915
            ScaleWidth      =   1035
            TabIndex        =   39
            Top             =   2160
            Width           =   1095
         End
         Begin VB.PictureBox Picture6 
            Height          =   975
            Left            =   1200
            Picture         =   "FRMCTTT.frx":A91A
            ScaleHeight     =   915
            ScaleWidth      =   1035
            TabIndex        =   38
            Top             =   1200
            Width           =   1095
         End
         Begin VB.PictureBox Picture5 
            Height          =   975
            Left            =   120
            Picture         =   "FRMCTTT.frx":E178
            ScaleHeight     =   915
            ScaleWidth      =   1035
            TabIndex        =   37
            Top             =   2160
            Width           =   1095
         End
         Begin VB.PictureBox Picture4 
            Height          =   975
            Left            =   1200
            Picture         =   "FRMCTTT.frx":119D6
            ScaleHeight     =   915
            ScaleWidth      =   1035
            TabIndex        =   36
            Top             =   2160
            Width           =   1095
         End
         Begin VB.PictureBox Picture3 
            Height          =   975
            Left            =   120
            Picture         =   "FRMCTTT.frx":15234
            ScaleHeight     =   915
            ScaleWidth      =   1035
            TabIndex        =   35
            Top             =   1200
            Width           =   1095
         End
         Begin VB.PictureBox Picture2 
            Height          =   975
            Left            =   1200
            Picture         =   "FRMCTTT.frx":18A92
            ScaleHeight     =   915
            ScaleWidth      =   1035
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Left            =   2280
            Picture         =   "FRMCTTT.frx":1C2F0
            ScaleHeight     =   915
            ScaleWidth      =   1035
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0000FFFF&
            Caption         =   "PLAY COMPUTER"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   3240
            Width           =   3015
         End
         Begin VB.CommandButton CMD9 
            BackColor       =   &H0000FFFF&
            Enabled         =   0   'False
            Height          =   735
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2280
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton CMD8 
            BackColor       =   &H00FFFF00&
            Enabled         =   0   'False
            Height          =   735
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2280
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton CMD7 
            BackColor       =   &H0000FFFF&
            Enabled         =   0   'False
            Height          =   735
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2280
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton CMD6 
            BackColor       =   &H008080FF&
            Enabled         =   0   'False
            Height          =   735
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton CMD5 
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   735
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton CMD4 
            BackColor       =   &H000000FF&
            Enabled         =   0   'False
            Height          =   735
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton CMD3 
            BackColor       =   &H00C000C0&
            Enabled         =   0   'False
            Height          =   735
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton CMD2 
            BackColor       =   &H0000FFFF&
            Enabled         =   0   'False
            Height          =   735
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton CMD1 
            BackColor       =   &H00FF00FF&
            Enabled         =   0   'False
            Height          =   735
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Line Line8 
            BorderWidth     =   3
            X1              =   2280
            X2              =   2280
            Y1              =   240
            Y2              =   3120
         End
         Begin VB.Line Line7 
            BorderWidth     =   3
            X1              =   1200
            X2              =   1200
            Y1              =   240
            Y2              =   3120
         End
         Begin VB.Line Line6 
            BorderWidth     =   5
            X1              =   3360
            X2              =   3360
            Y1              =   240
            Y2              =   3120
         End
         Begin VB.Line Line5 
            BorderWidth     =   5
            X1              =   120
            X2              =   3360
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Line Line4 
            BorderWidth     =   5
            X1              =   3360
            X2              =   120
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line3 
            BorderWidth     =   5
            X1              =   120
            X2              =   120
            Y1              =   240
            Y2              =   3120
         End
         Begin VB.Line Line2 
            BorderWidth     =   3
            X1              =   120
            X2              =   3360
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   120
            X2              =   3360
            Y1              =   1200
            Y2              =   1200
         End
      End
      Begin VB.OLE OLE1 
         Height          =   3735
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmcttt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer
Dim z As Integer
Private Sub CMD1_Click()
If A Mod 2 = 0 Then
A = A + 1
CMD1.Caption = "X"
CMD1.Enabled = False
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = False
CMD8.Enabled = False
CMD9.Enabled = False
End If
If CMD1.Caption = "X" And CMD2.Caption = "X" And CMD3.Caption = "X" Then
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
A = 0
frmcttt.Height = 5900
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD1.Caption = "X" And CMD4.Caption = "X" And CMD7.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
A = 0
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD1.Caption = "X" And CMD5.Caption = "X" And CMD9.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
A = 0
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub CMD2_Click()
If A Mod 2 = 0 Then
A = A + 1
CMD2.Caption = "X"
CMD2.Enabled = False
CMD1.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = False
CMD8.Enabled = False
CMD9.Enabled = False
End If

If CMD1.Caption = "X" And CMD2.Caption = "X" And CMD3.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
A = 0
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD5.Caption = "X" And CMD2.Caption = "X" And CMD8.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
A = 0
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub CMD3_Click()
If A Mod 2 = 0 Then
A = A + 1
CMD3.Caption = "X"
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = False
CMD8.Enabled = False
CMD9.Enabled = False
End If

If CMD1.Caption = "X" And CMD2.Caption = "X" And CMD3.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
A = 0
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD6.Caption = "X" And CMD9.Caption = "X" And CMD3.Caption = "X" Then
frmcttt.Height = 5900
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
A = 0
    CMD9.Enabled = False
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
End If
If CMD7.Caption = "X" And CMD5.Caption = "X" And CMD3.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
A = 0
    CMD9.Enabled = False
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub CMD4_Click()
If A Mod 2 = 0 Then
A = A + 1
CMD4.Caption = "X"
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = False
CMD8.Enabled = False
CMD9.Enabled = False
End If

If CMD1.Caption = "X" And CMD4.Caption = "X" And CMD7.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
A = 0
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD6.Caption = "X" And CMD4.Caption = "X" And CMD5.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
A = 0
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub CMD5_Click()
If A Mod 2 = 0 Then
A = A + 1
CMD5.Caption = "X"
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = False
CMD8.Enabled = False
CMD9.Enabled = False
End If

If CMD2.Caption = "X" And CMD5.Caption = "X" And CMD8.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
A = 0
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD6.Caption = "X" And CMD5.Caption = "X" And CMD4.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
A = 0
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD7.Caption = "X" And CMD5.Caption = "X" And CMD3.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
A = 0
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub CMD6_Click()
If A Mod 2 = 0 Then
A = A + 1
CMD6.Caption = "X"
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = False
CMD8.Enabled = False
CMD9.Enabled = False
End If

If CMD4.Caption = "X" And CMD5.Caption = "X" And CMD6.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
A = 0
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD6.Caption = "X" And CMD9.Caption = "X" And CMD3.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
A = 0
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub CMD7_Click()
If A Mod 2 = 0 Then
A = A + 1
CMD7.Caption = "X"
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = False
CMD8.Enabled = False
CMD9.Enabled = False
End If

If CMD1.Caption = "X" And CMD4.Caption = "X" And CMD7.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
A = 0
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD7.Caption = "X" And CMD9.Caption = "X" And CMD8.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
A = 0
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD7.Caption = "X" And CMD5.Caption = "X" And CMD3.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
A = 0
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub CMD8_Click()
If A Mod 2 = 0 Then
A = A + 1
CMD8.Caption = "X"
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = False
CMD8.Enabled = False
CMD9.Enabled = False
End If

If CMD5.Caption = "X" And CMD2.Caption = "X" And CMD8.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
A = 0
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD7.Caption = "X" And CMD8.Caption = "X" And CMD9.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
A = 0
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub CMD9_Click()
If A Mod 2 = 0 Then
A = A + 1
CMD9.Caption = "X"
CMD2.Enabled = False
CMD3.Enabled = False
CMD4.Enabled = False
CMD5.Enabled = False
CMD6.Enabled = False
CMD7.Enabled = False
CMD8.Enabled = False
CMD9.Enabled = False
End If

If CMD1.Caption = "X" And CMD5.Caption = "X" And CMD9.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
A = 0
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD6.Caption = "X" And CMD9.Caption = "X" And CMD3.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
A = 0
    CMD9.Enabled = False
End If
If CMD7.Caption = "X" And CMD8.Caption = "X" And CMD9.Caption = "X" Then
frmcttt.Height = 5900
LBL1.Caption = "CONGRADULATIONS " + TXTNAME.Text & Chr(13) + "YOU HAS WON THE GAME" + Chr(13) + "IT WAS NICE TO PLAY WITH YOU"
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
A = 0
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub Command1_Click()
If CMD7.Caption = "X" And CMD5.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD3.Caption = "" Then
        A = A + 1
        CMD3.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD7.Caption = "X" And CMD3.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD5.Caption = "" Then
        A = A + 1
        CMD5.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD3.Caption = "X" And CMD5.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD7.Caption = "" Then
        A = A + 1
        CMD7.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD5.Caption = "X" And CMD9.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD1.Caption = "" Then
        A = A + 1
        CMD1.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD1.Caption = "X" And CMD9.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD5.Caption = "" Then
        A = A + 1
        CMD5.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD1.Caption = "X" And CMD5.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD9.Caption = "" Then
        A = A + 1
        CMD9.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD8.Caption = "X" And CMD9.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD7.Caption = "" Then
        A = A + 1
        CMD7.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD7.Caption = "X" And CMD9.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD8.Caption = "" Then
        A = A + 1
        CMD8.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD7.Caption = "X" And CMD8.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD9.Caption = "" Then
        A = A + 1
        CMD9.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD5.Caption = "X" And CMD6.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD4.Caption = "" Then
        A = A + 1
        CMD4.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD6.Caption = "X" And CMD4.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD5.Caption = "" Then
        A = A + 1
        CMD5.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD5.Caption = "X" And CMD4.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD6.Caption = "" Then
        A = A + 1
        CMD6.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD3.Caption = "X" And CMD2.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD1.Caption = "" Then
        A = A + 1
        CMD1.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD1.Caption = "X" And CMD3.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD2.Caption = "" Then
        A = A + 1
        CMD2.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD1.Caption = "X" And CMD2.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD3.Caption = "" Then
        A = A + 1
        CMD3.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD6.Caption = "X" And CMD9.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD3.Caption = "" Then
        A = A + 1
        CMD3.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD3.Caption = "X" And CMD9.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD6.Caption = "" Then
        A = A + 1
        CMD6.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD3.Caption = "X" And CMD6.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD9.Caption = "" Then
        A = A + 1
        CMD9.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD8.Caption = "X" And CMD2.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD5.Caption = "" Then
        A = A + 1
        CMD5.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD5.Caption = "X" And CMD8.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD2.Caption = "" Then
        A = A + 1
        CMD2.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD2.Caption = "X" And CMD5.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD8.Caption = "" Then
        A = A + 1
        CMD8.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD1.Caption = "X" And CMD4.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD7.Caption = "" Then
        A = A + 1
        CMD7.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If
If CMD7.Caption = "X" And CMD1.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD4.Caption = "" Then
        A = A + 1
        CMD4.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If

If CMD7.Caption = "X" And CMD4.Caption = "X" Then
    If A Mod 2 = 1 Then
        If CMD1.Caption = "" Then
        A = A + 1
        CMD1.Caption = "O"
        End If
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    End If
End If

If A Mod 2 = 1 Then
    If CMD1.Caption = "" Then
    CMD1.Caption = "O"
    CMD1.Enabled = False
    A = A + 1
    End If
End If
If A Mod 2 = 1 Then
    If CMD2.Caption = "" Then
    CMD2.Caption = "O"
    CMD2.Enabled = False
    A = A + 1
    End If
End If
If A Mod 2 = 1 Then
    If CMD3.Caption = "" Then
    CMD3.Caption = "O"
    CMD3.Enabled = False
    A = A + 1
    End If
End If
If A Mod 2 = 1 Then
    If CMD4.Caption = "" Then
    CMD4.Caption = "O"
    CMD4.Enabled = False
    A = A + 1
    End If
End If
If A Mod 2 = 1 Then
    If CMD6.Caption = "" Then
    CMD6.Caption = "O"
    CMD6.Enabled = False
    A = A + 1
    End If
End If
If A Mod 2 = 1 Then
    If CMD5.Caption = "" Then
    CMD5.Caption = "O"
    CMD5.Enabled = False
    A = A + 1
    End If
End If
If A Mod 2 = 1 Then
    If CMD7.Caption = "" Then
    CMD7.Caption = "O"
    CMD7.Enabled = False
    A = A + 1
    End If
End If
If A Mod 2 = 1 Then
    If CMD8.Caption = "" Then
    CMD8.Caption = "O"
    CMD8.Enabled = False
    A = A + 1
    End If
End If
If A Mod 2 = 1 Then
    If CMD9.Caption = "" Then
    CMD9.Caption = "O"
    CMD9.Enabled = False
    A = A + 1
    End If
End If
If CMD1.Caption = "" Then
CMD1.Enabled = True
End If
If CMD2.Caption = "" Then
CMD2.Enabled = True
End If
If CMD3.Caption = "" Then
CMD3.Enabled = True
End If
If CMD4.Caption = "" Then
CMD4.Enabled = True
End If
If CMD5.Caption = "" Then
CMD5.Enabled = True
End If
If CMD6.Caption = "" Then
CMD6.Enabled = True
End If
If CMD7.Caption = "" Then
CMD7.Enabled = True
End If
If CMD8.Caption = "" Then
CMD8.Enabled = True
End If
If CMD9.Caption = "" Then
CMD9.Enabled = True
End If
If CMD1.Caption = "O" And CMD2.Caption = "O" And CMD3.Caption = "O" Then
frmcttt.Height = 5900
LBL1.Caption = "THE COMPUTER HAS WON THE GAME."
CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
A = 0
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD4.Caption = "O" And CMD5.Caption = "O" And CMD6.Caption = "O" Then
frmcttt.Height = 5900
LBL1.Caption = "THE COMPUTER HAS WON THE GAME."
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
A = 0
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD7.Caption = "O" And CMD8.Caption = "O" And CMD9.Caption = "O" Then
frmcttt.Height = 5900
LBL1.Caption = "THE COMPUTER HAS WON THE GAME."
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
A = 0
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD1.Caption = "O" And CMD4.Caption = "O" And CMD7.Caption = "O" Then
frmcttt.Height = 5900
LBL1.Caption = "THE COMPUTER HAS WON THE GAME."
    CMD2.Enabled = False
    CMD3.Enabled = False
A = 0
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD8.Caption = "O" And CMD2.Caption = "O" And CMD5.Caption = "O" Then
frmcttt.Height = 5900
LBL1.Caption = "THE COMPUTER HAS WON THE GAME."
    CMD2.Enabled = False
A = 0
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD6.Caption = "O" And CMD9.Caption = "O" And CMD3.Caption = "O" Then
frmcttt.Height = 5900
LBL1.Caption = "THE COMPUTER HAS WON THE GAME."
    CMD2.Enabled = False
    CMD3.Enabled = False
A = 0
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD1.Caption = "O" And CMD5.Caption = "O" And CMD9.Caption = "O" Then
frmcttt.Height = 5900
LBL1.Caption = "THE COMPUTER HAS WON THE GAME."
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
A = 0
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If CMD7.Caption = "O" And CMD5.Caption = "O" And CMD3.Caption = "O" Then
frmcttt.Height = 5900
LBL1.Caption = "THE COMPUTER HAS WON THE GAME."
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
A = 0
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
End If
If A = 9 Then
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End If
End Sub

Private Sub Command10_Click()
CMD1.Caption = ""
CMD2.Caption = ""
CMD3.Caption = ""
CMD4.Caption = ""
CMD5.Caption = ""
CMD6.Caption = ""
CMD7.Caption = ""
CMD8.Caption = ""
CMD9.Caption = ""
frmcttt.Height = 4725
Command10.Enabled = False
Command1.Enabled = False
TXTNAME.Text = ""
TXTAGE.Text = ""
Command12.Enabled = False
Command2.Enabled = False
Command13.Enabled = False
Command11.Enabled = False
CMD1.Enabled = False
CMD1.Visible = False
CMD2.Enabled = False
CMD2.Visible = False
CMD3.Enabled = False
CMD3.Visible = False
CMD4.Enabled = False
CMD4.Visible = False
CMD5.Enabled = False
CMD5.Visible = False
CMD6.Enabled = False
CMD6.Visible = False
CMD7.Enabled = False
CMD7.Visible = False
CMD8.Enabled = False
CMD8.Visible = False
CMD9.Enabled = False
CMD9.Visible = False
TXTNAME.SetFocus
A = 0
End Sub

Private Sub Command11_Click()
LBL1.Caption = ""
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
Command2.Enabled = True
Frame7.Enabled = True
Frame6.Enabled = False
CMD1.Enabled = True
CMD2.Enabled = True
CMD3.Enabled = True
CMD4.Enabled = True
CMD5.Enabled = True
CMD6.Enabled = True
CMD7.Enabled = True
CMD8.Enabled = True
CMD9.Enabled = True
CMD1.Visible = True
CMD2.Visible = True
CMD3.Visible = True
CMD4.Visible = True
CMD6.Visible = True
CMD5.Visible = True
CMD7.Visible = True
CMD8.Visible = True
CMD9.Visible = True
Command11.Enabled = False
Command1.Enabled = True
Command10.Enabled = True
Command12.Enabled = True
Frame2.Enabled = True
End Sub

Private Sub Command12_Click()
CMD1.Caption = ""
CMD2.Caption = ""
CMD3.Caption = ""
CMD4.Caption = ""
CMD5.Caption = ""
CMD6.Caption = ""
CMD7.Caption = ""
CMD8.Caption = ""
CMD9.Caption = ""
Frame6.Enabled = True
frmcttt.Height = 4725
Command13.Enabled = False
Command10.Enabled = False
Command1.Enabled = False
Command11.Enabled = True
Command12.Enabled = False
Command2.Enabled = False
CMD1.Enabled = False
CMD1.Visible = False
CMD2.Enabled = False
CMD2.Visible = False
CMD3.Enabled = False
CMD3.Visible = False
CMD4.Enabled = False
CMD4.Visible = False
CMD5.Enabled = False
CMD5.Visible = False
CMD6.Enabled = False
CMD6.Visible = False
CMD7.Enabled = False
CMD7.Visible = False
A = 0
CMD8.Enabled = False
CMD8.Visible = False
CMD9.Enabled = False
CMD9.Visible = False
End Sub

Private Sub Command13_Click()
If TXTNAME.Text = "subhanir" Or TXTNAME.Text = "Subhanir" Or TXTNAME.Text = "sUBHANIR" Or TXTNAME.Text = "SUBHANIR" Then
    If TXTAGE.Text = "INDIA" Then
    LBL1.Caption = "WELCOME MY CREATOR." + Chr(13) + "WOULD YOU LIKE TO PLAY A GAME WITH ME" + Chr(13) + "I KNOW, YOU ARE THE MASTER AND ME, THE COMPUTER,YOUR CREATION, YOUR SLAVE, CANNOT WIN AGAINST YOU."
    frmcttt.Height = 5900
    LBL1.Visible = True
    End If
End If
Command13.Enabled = False
Frame6.Enabled = True
Command11.Enabled = True
End Sub

Private Sub Command14_Click()
End
End Sub

Private Sub Command2_Click()
If z Mod 2 = 0 Then
frmcttt.Height = 5900
LBL1.Visible = True
z = z + 1
Else
frmcttt.Height = 4725

z = z + 1
End If

End Sub

Private Sub Form_Load()
A = 0
z = 0
End Sub

Private Sub TXTAGE_Change()
If TXTAGE.Text = "INDIA" Then
Command13.Enabled = True
End If
If TXTNAME.Text <> "" Then
If Val(TXTAGE.Text) > 5 Then
If Val(TXTAGE.Text) < 100 Then
Command13.Enabled = True
Else
frmcttt.Height = 5900
LBL1.Caption = "AGE GREATER THAN EXPECTED"
End If
Else
frmcttt.Height = 5900
LBL1.Caption = "AGE LESS THAN EXPECTED"
End If
Else
frmcttt.Height = 5900
LBL1.Caption = "MUST ENTER NAME TO CONTINUE"
End If
If TXTNAME.Text <> "" Then
If Val(TXTAGE.Text) > 5 Then
If Val(TXTAGE.Text) < 100 Then
frmcttt.Height = 5900
LBL1.Caption = "AGES EXCEPTED"
End If
frmcttt.Height = 5900
LBL1.Caption = "AGES EXCEPTED"
End If
frmcttt.Height = 5900
LBL1.Caption = "NAME EXCEPTED"
End If

End Sub

