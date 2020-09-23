VERSION 5.00
Begin VB.Form wizardbook 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Lord Callubonn's Wizards Spellbook Roller"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   9030
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Wizardsbook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   602
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List36 
      Height          =   405
      ItemData        =   "Wizardsbook.frx":0ECA
      Left            =   6480
      List            =   "Wizardsbook.frx":0ECC
      TabIndex        =   54
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List33 
      Height          =   405
      ItemData        =   "Wizardsbook.frx":0ECE
      Left            =   4680
      List            =   "Wizardsbook.frx":0ED0
      TabIndex        =   53
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List30 
      Height          =   405
      ItemData        =   "Wizardsbook.frx":0ED2
      Left            =   2880
      List            =   "Wizardsbook.frx":0ED4
      TabIndex        =   52
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List26 
      Height          =   405
      ItemData        =   "Wizardsbook.frx":0ED6
      Left            =   1200
      List            =   "Wizardsbook.frx":0ED8
      TabIndex        =   51
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List21 
      Height          =   405
      ItemData        =   "Wizardsbook.frx":0EDA
      Left            =   7560
      List            =   "Wizardsbook.frx":0EDC
      TabIndex        =   50
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List16 
      Height          =   405
      ItemData        =   "Wizardsbook.frx":0EDE
      Left            =   5760
      List            =   "Wizardsbook.frx":0EE0
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List11 
      Height          =   405
      ItemData        =   "Wizardsbook.frx":0EE2
      Left            =   3960
      List            =   "Wizardsbook.frx":0EE4
      TabIndex        =   48
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox Listf 
      Height          =   405
      ItemData        =   "Wizardsbook.frx":0EE6
      Left            =   2160
      List            =   "Wizardsbook.frx":0EE8
      TabIndex        =   47
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox Lista 
      Height          =   405
      ItemData        =   "Wizardsbook.frx":0EEA
      Left            =   360
      List            =   "Wizardsbook.frx":0EEC
      TabIndex        =   46
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text37 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   45
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text36 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   43
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text35 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   42
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text34 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   41
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text33 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   39
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text32 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text31 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text30 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   35
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text29 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox Text28 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text27 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text26 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text24 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text23 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text21 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 9th Level"
      Height          =   375
      Left            =   6480
      TabIndex        =   44
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 8th Level"
      Height          =   375
      Left            =   4680
      TabIndex        =   40
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 7th Level"
      Height          =   375
      Left            =   2880
      TabIndex        =   36
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 6th Level"
      Height          =   375
      Left            =   1080
      TabIndex        =   31
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 5th Level"
      Height          =   375
      Left            =   7440
      TabIndex        =   25
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 4th Level"
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 3rd Level"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 2nd Level"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 1st Level"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Picture         =   "Wizardsbook.frx":0EEE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
   Begin VB.Menu mnuNew 
      Caption         =   "&New"
      Begin VB.Menu mnu1st 
         Caption         =   "&1st Level Only"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu2nd 
         Caption         =   "&2nd Level Only"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu3rd 
         Caption         =   "&3rd Level Only"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu4th 
         Caption         =   "&4th Level Only"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnu5th 
         Caption         =   "&5th Level Only"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu6th 
         Caption         =   "&6th Level Only"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu7th 
         Caption         =   "&7th Level Only"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnu8th 
         Caption         =   "&8th Level Only"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnu9th 
         Caption         =   "&9th Level Only"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "wizardbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

FreeFileHandle = FreeFile

Open App.Path & "\1st.txt" For Input As #FreeFileHandle  ' opens textfile to read

Do

' Inputs the whole line at once.

Line Input #FreeFileHandle, theline$

' This list box will act as an array, exept u won't have to declara how big it is...
' Listbox are almost infinite

  Lista.AddItem theline$

Loop While Not EOF(FreeFileHandle) ' EOF( End Of File)

' Ok, now thats it... All them are added to a list box..

Close #FreeFileHandle

'*

FreeFileHandle = FreeFile
Open App.Path & "\2nd.txt" For Input As #FreeFileHandle  ' opens textfile to read

Do

' Inputs the whole line at once.

Line Input #FreeFileHandle, theline$

' This list box will act as an array, exept u won't have to declara how big it is...
' Listbox are almost infinite

  Listf.AddItem theline$

Loop While Not EOF(FreeFileHandle) ' EOF( End Of File)

' Ok, now thats it... All them are added to a list box..

Close #FreeFileHandle

'*
FreeFileHandle = FreeFile

Open App.Path & "\3rd.txt" For Input As #FreeFileHandle  ' opens textfile to read

Do

' Inputs the whole line at once.

Line Input #FreeFileHandle, theline$

' This list box will act as an array, exept u won't have to declara how big it is...
' Listbox are almost infinite

  List11.AddItem theline$

Loop While Not EOF(FreeFileHandle) ' EOF( End Of File)

' Ok, now thats it... All them are added to a list box..

Close #FreeFileHandle

'*

FreeFileHandle = FreeFile

Open App.Path & "\4th.txt" For Input As #FreeFileHandle  ' opens textfile to read

Do

' Inputs the whole line at once.

Line Input #FreeFileHandle, theline$

' This list box will act as an array, exept u won't have to declara how big it is...
' Listbox are almost infinite

  List16.AddItem theline$

Loop While Not EOF(FreeFileHandle) ' EOF( End Of File)

' Ok, now thats it... All them are added to a list box..

Close #FreeFileHandle

'*

FreeFileHandle = FreeFile

Open App.Path & "\5th.txt" For Input As #FreeFileHandle  ' opens textfile to read

Do

' Inputs the whole line at once.

Line Input #FreeFileHandle, theline$

' This list box will act as an array, exept u won't have to declara how big it is...
' Listbox are almost infinite

  List21.AddItem theline$

Loop While Not EOF(FreeFileHandle) ' EOF( End Of File)

' Ok, now thats it... All them are added to a list box..

Close #FreeFileHandle

'*

FreeFileHandle = FreeFile

Open App.Path & "\6th.txt" For Input As #FreeFileHandle  ' opens textfile to read

Do

' Inputs the whole line at once.

Line Input #FreeFileHandle, theline$

' This list box will act as an array, exept u won't have to declara how big it is...
' Listbox are almost infinite

  List26.AddItem theline$

Loop While Not EOF(FreeFileHandle) ' EOF( End Of File)

' Ok, now thats it... All them are added to a list box..

Close #FreeFileHandle

'*

FreeFileHandle = FreeFile

Open App.Path & "\7th.txt" For Input As #FreeFileHandle  ' opens textfile to read

Do

' Inputs the whole line at once.

Line Input #FreeFileHandle, theline$

' This list box will act as an array, exept u won't have to declara how big it is...
' Listbox are almost infinite

  List30.AddItem theline$

Loop While Not EOF(FreeFileHandle) ' EOF( End Of File)

' Ok, now thats it... All them are added to a list box..

Close #FreeFileHandle

'*

FreeFileHandle = FreeFile

Open App.Path & "\8th.txt" For Input As #FreeFileHandle  ' opens textfile to read

Do

' Inputs the whole line at once.

Line Input #FreeFileHandle, theline$

' This list box will act as an array, exept u won't have to declara how big it is...
' Listbox are almost infinite

  List33.AddItem theline$

Loop While Not EOF(FreeFileHandle) ' EOF( End Of File)

' Ok, now thats it... All them are added to a list box..

Close #FreeFileHandle

'*
FreeFileHandle = FreeFile

Open App.Path & "\9th.txt" For Input As #FreeFileHandle  ' opens textfile to read

Do

' Inputs the whole line at once.

Line Input #FreeFileHandle, theline$

' This list box will act as an array, exept u won't have to declara how big it is...
' Listbox are almost infinite

  List36.AddItem theline$

Loop While Not EOF(FreeFileHandle) ' EOF( End Of File)

' Ok, now thats it... All them are added to a list box..

Close #FreeFileHandle

End Sub

Private Sub mnu1st_Click()

'  Maxnumber = The Number of entries on file  , so it should look like this
  
  Maxnumber = Lista.ListCount ' get how many items on the list
  luckynumber = Int(Rnd * Maxnumber) + 1
   luckynumber2 = Int(Rnd * Maxnumber) + 1
  luckynumber3 = Int(Rnd * Maxnumber) + 1
  luckynumber4 = Int(Rnd * Maxnumber) + 1
  luckynumber5 = Int(Rnd * Maxnumber) + 1

Text1.Text = Lista.List(luckynumber)
Text2.Text = Lista.List(luckynumber2)
Text3.Text = Lista.List(luckynumber3)
Text4.Text = Lista.List(luckynumber4)
Text5.Text = Lista.List(luckynumber5)

End Sub

Private Sub mnu2nd_Click()

'  Maxnumber = The Number of entries on file  , so it should look like this
  
  Maxnumber = Listf.ListCount ' get how many items on the list
  luckynumber = Int(Rnd * Maxnumber) + 1
 luckynumber2 = Int(Rnd * Maxnumber) + 1
  luckynumber3 = Int(Rnd * Maxnumber) + 1
  luckynumber4 = Int(Rnd * Maxnumber) + 1
  luckynumber5 = Int(Rnd * Maxnumber) + 1

Text6.Text = Listf.List(luckynumber)
Text7.Text = Listf.List(luckynumber2)
Text8.Text = Listf.List(luckynumber3)
Text9.Text = Listf.List(luckynumber4)
Text10.Text = Listf.List(luckynumber5)

End Sub

Private Sub mnu3rd_Click()

'  Maxnumber = The Number of entries on file  , so it should look like this
  
  Maxnumber = List11.ListCount ' get how many items on the list
  luckynumber = Int(Rnd * Maxnumber) + 1
 luckynumber2 = Int(Rnd * Maxnumber) + 1
  luckynumber3 = Int(Rnd * Maxnumber) + 1
  luckynumber4 = Int(Rnd * Maxnumber) + 1
  luckynumber5 = Int(Rnd * Maxnumber) + 1

Text11.Text = List11.List(luckynumber)
Text12.Text = List11.List(luckynumber2)
Text13.Text = List11.List(luckynumber3)
Text14.Text = List11.List(luckynumber4)
Text15.Text = List11.List(luckynumber5)

End Sub

Private Sub mnu4th_Click()

'  Maxnumber = The Number of entries on file  , so it should look like this
  
  Maxnumber = List16.ListCount ' get how many items on the list
  luckynumber = Int(Rnd * Maxnumber) + 1
 luckynumber2 = Int(Rnd * Maxnumber) + 1
  luckynumber3 = Int(Rnd * Maxnumber) + 1
  luckynumber4 = Int(Rnd * Maxnumber) + 1
  luckynumber5 = Int(Rnd * Maxnumber) + 1

Text16.Text = List16.List(luckynumber)
Text17.Text = List16.List(luckynumber2)
Text18.Text = List16.List(luckynumber3)
Text19.Text = List16.List(luckynumber4)
Text20.Text = List16.List(luckynumber5)

End Sub

Private Sub mnu5th_Click()

'  Maxnumber = The Number of entries on file  , so it should look like this
  
  Maxnumber = List21.ListCount ' get how many items on the list
  luckynumber = Int(Rnd * Maxnumber) + 1
 luckynumber2 = Int(Rnd * Maxnumber) + 1
  luckynumber3 = Int(Rnd * Maxnumber) + 1
  luckynumber4 = Int(Rnd * Maxnumber) + 1
  luckynumber5 = Int(Rnd * Maxnumber) + 1

Text21.Text = List21.List(luckynumber)
Text22.Text = List21.List(luckynumber2)
Text23.Text = List21.List(luckynumber3)
Text24.Text = List21.List(luckynumber4)
Text25.Text = List21.List(luckynumber5)

End Sub

Private Sub mnu6th_Click()

  Maxnumber = List26.ListCount ' get how many items on the list
  luckynumber = Int(Rnd * Maxnumber) + 1
 luckynumber2 = Int(Rnd * Maxnumber) + 1
  luckynumber3 = Int(Rnd * Maxnumber) + 1
  luckynumber4 = Int(Rnd * Maxnumber) + 1

Text26.Text = List26.List(luckynumber)
Text27.Text = List26.List(luckynumber2)
Text28.Text = List26.List(luckynumber3)
Text29.Text = List26.List(luckynumber4)

End Sub

Private Sub mnu7th_Click()

  Maxnumber = List30.ListCount ' get how many items on the list
  luckynumber = Int(Rnd * Maxnumber) + 1
 luckynumber2 = Int(Rnd * Maxnumber) + 1
  luckynumber3 = Int(Rnd * Maxnumber) + 1

Text30.Text = List30.List(luckynumber)
Text31.Text = List30.List(luckynumber2)
Text32.Text = List30.List(luckynumber3)

End Sub

Private Sub mnu8th_Click()
  
  Maxnumber = List33.ListCount ' get how many items on the list
  luckynumber = Int(Rnd * Maxnumber) + 1
 luckynumber2 = Int(Rnd * Maxnumber) + 1
  luckynumber3 = Int(Rnd * Maxnumber) + 1

Text33.Text = List33.List(luckynumber)
Text34.Text = List33.List(luckynumber2)
Text35.Text = List33.List(luckynumber3)

End Sub

Private Sub mnu9th_Click()
 
  Maxnumber = List36.ListCount ' get how many items on the list
  luckynumber = Int(Rnd * Maxnumber) + 1
 luckynumber2 = Int(Rnd * Maxnumber) + 1
 
Text36.Text = List36.List(luckynumber)
Text37.Text = List36.List(luckynumber2)

End Sub

Private Sub mnuAbout_Click()
Load about
about.Show
End Sub

Private Sub mnuPrint_Click()

On Error GoTo Errorhandler
'Image1.Visible = False
Me.PrintForm
Image1.Visible = True
Exit Sub

Errorhandler:
Exit Sub

End Sub
