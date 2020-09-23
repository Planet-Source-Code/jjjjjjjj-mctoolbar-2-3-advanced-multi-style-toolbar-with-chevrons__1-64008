VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " McToolBar 2.3 - Test Form !!"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Test.McToolBar McToolBar 
      Height          =   585
      Index           =   15
      Left            =   6120
      TabIndex        =   87
      ToolTipText     =   "Button Pressed"
      Top             =   1320
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   1032
      BackColor       =   8421504
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   11
      ButtonsWidth    =   77
      ButtonsPerRow   =   11
      HoverColor      =   8388608
      BackGradient    =   3
      ButtonsMode     =   6
      ButtonsBackColor=   12632256
      ButtonsGradient =   3
      ButtonsPerRow_Chev=   1
      ShowChevron     =   -1  'True
      ButtonToolTipText1=   "Open Files"
      ButtonToolTipIcon1=   1
      Button_Type1    =   1
      ButtonCaption2  =   "Open"
      ButtonIcon2     =   "frmTest.frx":000C
      ButtonToolTipText2=   "Open Files"
      ButtonToolTipIcon2=   1
      ButtonIconAllignment2=   2
      ButtonCaption3  =   "Save"
      ButtonIcon3     =   "frmTest.frx":03A6
      ButtonToolTipText3=   "Save Files"
      ButtonToolTipIcon3=   1
      ButtonIconAllignment3=   2
      ButtonCaption4  =   "Print"
      ButtonIcon4     =   "frmTest.frx":0740
      ButtonToolTipText4=   "Print preview"
      ButtonToolTipIcon4=   1
      ButtonIconAllignment4=   2
      Button_Type5    =   1
      ButtonCaption6  =   "Cut"
      ButtonIcon6     =   "frmTest.frx":0ADA
      ButtonToolTipText6=   "Cut now"
      ButtonIconAllignment6=   2
      ButtonCaption7  =   "Copy"
      ButtonIcon7     =   "frmTest.frx":0E74
      ButtonToolTipText7=   "Copy now"
      ButtonIconAllignment7=   2
      ButtonCaption8  =   "Paste"
      ButtonIcon8     =   "frmTest.frx":120E
      ButtonToolTipText8=   "Paste from ClipBoard...."
      ButtonToolTipIcon8=   1
      ButtonIconAllignment8=   2
      Button_Type9    =   1
      ButtonCaption10 =   "Previous"
      ButtonIcon10    =   "frmTest.frx":15A8
      ButtonPressed10 =   -1  'True
      ButtonIconAllignment10=   2
      ButtonCaption11 =   "Next"
      ButtonIcon11    =   "frmTest.frx":1942
      ButtonIconAllignment11=   3
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   2
      Left            =   6360
      ScaleHeight     =   4785
      ScaleWidth      =   4065
      TabIndex        =   16
      Top             =   2040
      Width           =   4095
      Begin Test.McToolBar McToolBar2 
         Height          =   4620
         Left            =   3120
         TabIndex        =   60
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   8149
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   7
         ButtonsWidth    =   63
         ButtonsHeight   =   44
         ButtonsPerRow   =   1
         HoverColor      =   16761087
         TooTipStyle     =   0
         ButtonsMode     =   4
         ButtonCaption1  =   "Flat"
         ButtonToolTipIcon1=   1
         ButtonCaption2  =   "Soft"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   "Solid"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   "Win 98"
         ButtonToolTipIcon4=   1
         ButtonCaption5  =   "Office XP"
         ButtonToolTipIcon5=   1
         ButtonCaption6  =   "XP"
         ButtonToolTipIcon6=   1
         ButtonCaption7  =   "Plastic"
         ButtonToolTipIcon7=   1
      End
      Begin Test.McToolBar McToolBar 
         Height          =   480
         Index           =   1
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Button Pressed"
         Top             =   120
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   8438015
         ButtonsMode     =   0
         ButtonsBackColor=   12648447
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":1CDC
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":2076
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":2410
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":27AA
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":2B44
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":2EDE
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":3278
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":3612
      End
      Begin Test.McToolBar McToolBar 
         Height          =   480
         Index           =   0
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Button Pressed"
         Top             =   720
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   8438015
         ButtonsMode     =   1
         ButtonsBackColor=   12648447
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":39AC
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":3D46
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":40E0
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":447A
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":4814
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":4BAE
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":4F48
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":52E2
      End
      Begin Test.McToolBar McToolBar 
         Height          =   480
         Index           =   2
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Button Pressed"
         Top             =   2760
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   16744576
         ButtonsMode     =   4
         ButtonsBackColor=   12648447
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":567C
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":5A16
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":5DB0
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":614A
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":64E4
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":687E
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":6C18
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":6FB2
      End
      Begin Test.McToolBar McToolBar 
         Height          =   585
         Index           =   3
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Button Pressed"
         Top             =   2040
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1032
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   8438015
         ButtonsBackColor=   12648447
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":734C
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":76E6
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":7A80
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":7E1A
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":81B4
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":854E
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":88E8
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":8C82
      End
      Begin Test.McToolBar McToolBar 
         Height          =   585
         Index           =   5
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Button Pressed"
         Top             =   4080
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1032
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   8388608
         ButtonsMode     =   6
         ButtonsBackColor=   -2147483644
         ButtonsGradientCol=   -2147483633
         ButtonsGradient =   3
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":901C
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":93B6
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":9750
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":9AEA
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":9E84
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":A21E
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":A5B8
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":A952
      End
      Begin Test.McToolBar McToolBar 
         Height          =   480
         Index           =   6
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Button Pressed"
         Top             =   3360
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   847
         BackColor       =   14807794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   8438015
         ButtonsMode     =   5
         ButtonsBackColor=   14807794
         ButtonsGradient =   3
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":ACEC
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":B086
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":B420
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":B7BA
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":BB54
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":BEEE
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":C288
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":C622
      End
      Begin Test.McToolBar McToolBar 
         Height          =   585
         Index           =   4
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Button Pressed"
         Top             =   1320
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1032
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   8438015
         ButtonsMode     =   2
         ButtonsBackColor=   12648447
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":C9BC
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":CD56
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":D0F0
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":D48A
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":D824
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":DBBE
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":DF58
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":E2F2
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   10605
      TabIndex        =   0
      Top             =   0
      Width           =   10605
      Begin VB.Image Image1 
         Height          =   495
         Left            =   240
         Picture         =   "frmTest.frx":E68C
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "McToolBar 2.3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "     An advanced XP style Toolbar (Single file'd) with Chevrons..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   5505
      End
   End
   Begin Test.McToolBar McToolBar1 
      Align           =   1  'Align Top
      Height          =   585
      Left            =   0
      TabIndex        =   84
      ToolTipText     =   "Button Pressed"
      Top             =   975
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   1032
      BackColor       =   12632256
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   11
      BackGround      =   "frmTest.frx":F7D2
      ButtonsPerRow   =   11
      HoverColor      =   8438015
      BackGradient    =   3
      ButtonsMode     =   2
      ButtonsBackColor=   12648447
      ButtonsPerRow_Chev=   11
      ButtonToolTipText1=   "Open Files"
      ButtonToolTipIcon1=   1
      Button_Type1    =   1
      ButtonCaption2  =   ""
      ButtonIcon2     =   "frmTest.frx":12524
      ButtonToolTipText2=   "Open Files"
      ButtonToolTipIcon2=   1
      ButtonCaption3  =   ""
      ButtonIcon3     =   "frmTest.frx":128BE
      ButtonToolTipText3=   "Save Files"
      ButtonToolTipIcon3=   1
      ButtonCaption4  =   ""
      ButtonIcon4     =   "frmTest.frx":12C58
      ButtonToolTipText4=   "Print preview"
      ButtonToolTipIcon4=   1
      Button_Type5    =   1
      ButtonCaption6  =   ""
      ButtonIcon6     =   "frmTest.frx":12FF2
      ButtonToolTipText6=   "Cut now"
      ButtonCaption7  =   ""
      ButtonIcon7     =   "frmTest.frx":1338C
      ButtonToolTipText7=   "Copy now"
      ButtonCaption8  =   ""
      ButtonIcon8     =   "frmTest.frx":13726
      ButtonToolTipText8=   "Paste Now"
      Button_Type9    =   1
      ButtonPressed10 =   -1  'True
      ButtonPressed11 =   -1  'True
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   1
      Left            =   240
      ScaleHeight     =   4785
      ScaleWidth      =   5865
      TabIndex        =   15
      Top             =   2040
      Width           =   5895
      Begin VB.CheckBox chkChev 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show Chevron"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   86
         Top             =   960
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.CheckBox chkSep 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show Seperator"
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
         Left            =   360
         TabIndex        =   85
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2040
      End
      Begin VB.ComboBox cmbType 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmTest.frx":13AC0
         Left            =   3360
         List            =   "frmTest.frx":13ACA
         TabIndex        =   82
         Text            =   "    [TYP_Button] = 0"
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtRowChev 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   44
         Text            =   "1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtindex 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   34
         Text            =   "1"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtCaption 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   33
         Text            =   "Open"
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtRow 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   32
         Text            =   "11"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtTooltip 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   31
         Text            =   "Open files"
         Top             =   3000
         Width           =   2175
      End
      Begin VB.ComboBox cmbIcon 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmTest.frx":13AFD
         Left            =   1080
         List            =   "frmTest.frx":13B0D
         TabIndex        =   30
         Text            =   "[Icon_None] = 0"
         Top             =   3480
         Width           =   2175
      End
      Begin VB.ComboBox cmbStyle 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmTest.frx":13B59
         Left            =   1080
         List            =   "frmTest.frx":13B63
         TabIndex        =   29
         Text            =   "[Tip_Normal] = 1"
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   28
         Text            =   "32"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   27
         Text            =   "32"
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkEnabled 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   26
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkEnablectl 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enabled"
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
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin VB.CheckBox chkPress 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pressed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   24
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox cmbCapAln 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmTest.frx":13B8C
         Left            =   3360
         List            =   "frmTest.frx":13B9F
         TabIndex        =   23
         Text            =   "    [ALN_Center] = 4"
         Top             =   3480
         Width           =   2175
      End
      Begin Test.McToolBar tbApply 
         Height          =   375
         Left            =   3720
         TabIndex        =   48
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonsWidth    =   99
         ButtonsHeight   =   25
         ButtonsPerRow   =   11
         HoverColor      =   8388608
         TooTipStyle     =   0
         ButtonsMode     =   6
         ButtonsBackColor=   -2147483644
         ButtonsGradientCol=   14737632
         ButtonsGradient =   3
         ButtonCaption1  =   "Apply Changes"
         ButtonToolTipIcon1=   1
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Button Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   83
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buttons per Row Chevron"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3600
         TabIndex        =   45
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Button Index"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   360
         TabIndex        =   43
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   42
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buttons per Row"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3600
         TabIndex        =   41
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "ToolTip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   40
         Top             =   3120
         Width           =   510
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tool Icon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   39
         Top             =   3480
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H80000010&
         Height          =   1455
         Left            =   120
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buttons Height"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   38
         Top             =   360
         Width           =   540
      End
      Begin VB.Label sss 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buttons Width"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   37
         Top             =   360
         Width           =   540
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H80000010&
         Height          =   3015
         Left            =   120
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tool Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   36
         Top             =   3960
         Width           =   540
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Icon Allignment!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   35
         Top             =   3240
         Width           =   1380
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   0
      Left            =   240
      ScaleHeight     =   4785
      ScaleWidth      =   5865
      TabIndex        =   3
      Top             =   2040
      Width           =   5895
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000010&
         Height          =   4575
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   5535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   5280
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Give the new index to the property ""ButtonMove"". Selected Button will move to the given index"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   14
         Top             =   4200
         Width           =   5055
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Move Button :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   13
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "In property window set ""ButtonRemove"" to ""Yes!"". The selected button will be removed!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   3480
         Width           =   5055
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remove Button :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   11
         Top             =   3240
         Width           =   1440
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "In property window set ""Button_Count"". This much buttons will be created instantly with default values"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create Button :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTest.frx":13C0B
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   8
         Top             =   2520
         Width           =   5055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assign Properties :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1680
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "In property window set the ""Button_Index"" (is shown on the control, underlined in design time )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "          Asign all the properties at design time without using property window! [ Read the following carefully! ]"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Button :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1290
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   3
      Left            =   240
      ScaleHeight     =   4905
      ScaleWidth      =   5865
      TabIndex        =   50
      Top             =   2040
      Width           =   5895
      Begin Test.McToolBar McToolBar 
         Height          =   480
         Index           =   7
         Left            =   360
         TabIndex        =   68
         ToolTipText     =   "Button Pressed"
         Top             =   600
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   16744576
         BackGradient    =   3
         ButtonsMode     =   4
         ButtonsBackColor=   12648447
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":13CBD
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":14057
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":143F1
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":1478B
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":14B25
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":14EBF
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":15259
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":155F3
      End
      Begin Test.McToolBar McToolBar 
         Height          =   585
         Index           =   8
         Left            =   360
         TabIndex        =   70
         ToolTipText     =   "Button Pressed"
         Top             =   1560
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   1032
         BackColor       =   8421504
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   8388608
         BackGradient    =   3
         ButtonsMode     =   6
         ButtonsBackColor=   12632256
         ButtonsGradient =   3
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":1598D
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":15D27
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":160C1
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":1645B
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":167F5
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":16B8F
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":16F29
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":172C3
      End
      Begin Test.McToolBar McToolBar 
         Height          =   585
         Index           =   9
         Left            =   3240
         TabIndex        =   72
         ToolTipText     =   "Button Pressed"
         Top             =   3840
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   1032
         BackColor       =   8421631
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   65280
         BackGradient    =   3
         ButtonsMode     =   6
         ButtonsBackColor=   192
         ButtonsGradient =   3
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":1765D
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":179F7
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":17D91
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":1812B
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":184C5
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":1885F
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":18BF9
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":18F93
      End
      Begin Test.McToolBar McToolBar 
         Height          =   480
         Index           =   10
         Left            =   2880
         TabIndex        =   74
         ToolTipText     =   "Button Pressed"
         Top             =   600
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         ButtonsPerRow   =   11
         HoverColor      =   8438015
         BackGradient    =   3
         ButtonsMode     =   0
         ButtonsBackColor=   12648447
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":1932D
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":196C7
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":19A61
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":19DFB
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":1A195
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":1A52F
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":1A8C9
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":1AC63
      End
      Begin Test.McToolBar McToolBar 
         Height          =   585
         Index           =   11
         Left            =   360
         TabIndex        =   76
         ToolTipText     =   "Button Pressed"
         Top             =   2640
         Width           =   2280
         _ExtentX        =   5292
         _ExtentY        =   1032
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   11
         BackGround      =   "frmTest.frx":1AFFD
         ButtonsPerRow   =   11
         HoverColor      =   8438015
         ButtonsMode     =   2
         ButtonsBackColor=   12648447
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         Button_Type1    =   1
         ButtonCaption2  =   ""
         ButtonIcon2     =   "frmTest.frx":1DD4F
         ButtonToolTipText2=   "Open Files"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   ""
         ButtonIcon3     =   "frmTest.frx":1E0E9
         ButtonToolTipText3=   "Save Files"
         ButtonToolTipIcon3=   1
         ButtonCaption4  =   ""
         ButtonIcon4     =   "frmTest.frx":1E483
         ButtonToolTipText4=   "Print preview"
         ButtonToolTipIcon4=   1
         Button_Type5    =   1
         ButtonCaption6  =   ""
         ButtonIcon6     =   "frmTest.frx":1E81D
         ButtonToolTipText6=   "Cut now"
         ButtonCaption7  =   ""
         ButtonIcon7     =   "frmTest.frx":1EBB7
         ButtonToolTipText7=   "Copy now"
         ButtonCaption8  =   ""
         ButtonIcon8     =   "frmTest.frx":1EF51
         ButtonToolTipText8=   "Paste Now"
         Button_Type9    =   1
         ButtonCaption10 =   ""
         ButtonIcon10    =   "frmTest.frx":1F2EB
         ButtonPressed10 =   -1  'True
         ButtonCaption11 =   ""
         ButtonIcon11    =   "frmTest.frx":1F685
      End
      Begin Test.McToolBar McToolBar 
         Height          =   480
         Index           =   12
         Left            =   360
         TabIndex        =   78
         ToolTipText     =   "Button Pressed"
         Top             =   3960
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   847
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         ButtonsWidth    =   60
         ButtonsPerRow   =   11
         HoverColor      =   8388608
         ButtonsMode     =   6
         ShowSeperator   =   0   'False
         ButtonsBackColor=   -2147483644
         ButtonsGradientCol=   -2147483633
         ButtonsGradient =   3
         ButtonsPerRow_Chev=   1
         ShowChevron     =   -1  'True
         ButtonCaption1  =   "Cancel"
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         ButtonCaption2  =   "Ok"
         ButtonToolTipIcon2=   1
         ButtonCaption3  =   "Apply"
         ButtonToolTipIcon3=   1
      End
      Begin Test.McToolBar McToolBar 
         Height          =   1440
         Index           =   13
         Left            =   2880
         TabIndex        =   80
         ToolTipText     =   "Button Pressed"
         Top             =   1680
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   2540
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         ButtonsWidth    =   77
         ButtonsPerRow   =   1
         HoverColor      =   12648384
         ButtonsMode     =   5
         ButtonsBackColor=   12632319
         ButtonsGradient =   3
         ButtonsPerRow_Chev=   1
         ButtonCaption1  =   "Open"
         ButtonIcon1     =   "frmTest.frx":1FA1F
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         ButtonIconAllignment1=   2
         ButtonCaption2  =   "Save"
         ButtonIcon2     =   "frmTest.frx":1FDB9
         ButtonToolTipText2=   "Save Files"
         ButtonToolTipIcon2=   1
         ButtonIconAllignment2=   2
         ButtonCaption3  =   "Print"
         ButtonIcon3     =   "frmTest.frx":20153
         ButtonToolTipText3=   "Print preview"
         ButtonToolTipIcon3=   1
         ButtonIconAllignment3=   2
      End
      Begin Test.McToolBar McToolBar 
         Height          =   2580
         Index           =   14
         Left            =   4200
         TabIndex        =   81
         ToolTipText     =   "Button Pressed"
         Top             =   1200
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   4551
         BackColor       =   16761024
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Button_Count    =   3
         ButtonsWidth    =   77
         ButtonsHeight   =   55
         ButtonsPerRow   =   1
         HoverColor      =   16744576
         BackGradient    =   1
         ButtonsMode     =   4
         ButtonsBackColor=   16761024
         ButtonsPerRow_Chev=   1
         ButtonCaption1  =   "Open"
         ButtonIcon1     =   "frmTest.frx":204ED
         ButtonToolTipText1=   "Open Files"
         ButtonToolTipIcon1=   1
         ButtonIconAllignment1=   0
         ButtonCaption2  =   "Save"
         ButtonIcon2     =   "frmTest.frx":20887
         ButtonToolTipText2=   "Save Files"
         ButtonToolTipIcon2=   1
         ButtonIconAllignment2=   0
         ButtonCaption3  =   "Print"
         ButtonIcon3     =   "frmTest.frx":20C21
         ButtonToolTipText3=   "Print preview"
         ButtonToolTipIcon3=   1
         ButtonIconAllignment3=   0
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I am smart"
         Height          =   195
         Left            =   360
         TabIndex        =   79
         Top             =   3600
         Width           =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bitmap tiled"
         Height          =   195
         Left            =   360
         TabIndex        =   77
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flat with gradient"
         Height          =   195
         Left            =   2880
         TabIndex        =   75
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAD "
         Height          =   195
         Left            =   3360
         TabIndex        =   73
         Top             =   3600
         Width           =   405
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plastic - The dark side"
         Height          =   195
         Left            =   360
         TabIndex        =   71
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OfficeXP with gradient"
         Height          =   195
         Left            =   360
         TabIndex        =   69
         Top             =   360
         Width           =   1575
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H80000010&
         Height          =   4575
         Left            =   120
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   2
      Left            =   240
      ScaleHeight     =   4905
      ScaleWidth      =   5865
      TabIndex        =   47
      Top             =   2040
      Width           =   5895
      Begin VB.ComboBox cmbBorder 
         Height          =   315
         ItemData        =   "frmTest.frx":20FBB
         Left            =   360
         List            =   "frmTest.frx":20FC8
         TabIndex        =   65
         Text            =   "    BDR_None = 0"
         Top             =   4080
         Width           =   2175
      End
      Begin VB.ComboBox cmbBtnGrd 
         Height          =   315
         ItemData        =   "frmTest.frx":2100C
         Left            =   360
         List            =   "frmTest.frx":21025
         TabIndex        =   63
         Text            =   "[Fill_None] = 0"
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optBtnGrad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Btn Back Gradient Col"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   62
         Top             =   2160
         Width           =   2055
      End
      Begin VB.OptionButton optBtnBack 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Btn BackColor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   61
         Top             =   1920
         Width           =   1575
      End
      Begin VB.PictureBox picCol 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   3360
         Picture         =   "frmTest.frx":210DC
         ScaleHeight     =   2100
         ScaleWidth      =   2100
         TabIndex        =   58
         Top             =   1080
         Width           =   2130
      End
      Begin VB.OptionButton optBack 
         BackColor       =   &H00E0E0E0&
         Caption         =   "BackColor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   57
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton optHover 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hover Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   56
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton optFore 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fore Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   55
         Top             =   2760
         Width           =   1575
      End
      Begin VB.OptionButton optBackGrd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Back Gradient Col"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   54
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton optTipBack 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ToolTip backCol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   53
         Top             =   3120
         Width           =   1935
      End
      Begin VB.OptionButton optTipFore 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ToolTip BackCol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   52
         Top             =   3360
         Width           =   2055
      End
      Begin VB.ComboBox cmbGradient 
         Height          =   315
         ItemData        =   "frmTest.frx":2F6CE
         Left            =   3360
         List            =   "frmTest.frx":2F6E7
         TabIndex        =   51
         Text            =   "[Fill_None] = 0"
         Top             =   600
         Width           =   2175
      End
      Begin Test.McToolBar tbTile 
         Height          =   375
         Left            =   3360
         TabIndex        =   67
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonsWidth    =   99
         ButtonsHeight   =   25
         ButtonsPerRow   =   11
         HoverColor      =   8388608
         TooTipStyle     =   0
         ButtonsMode     =   6
         ButtonsBackColor=   -2147483644
         ButtonsGradientCol=   14737632
         ButtonsGradient =   3
         ButtonCaption1  =   "Apply Tile"
         ButtonToolTipIcon1=   1
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Border Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   66
         Top             =   3840
         Width           =   885
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Btn BackGradient"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   64
         Top             =   360
         Width           =   1230
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000010&
         Height          =   2655
         Left            =   240
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "BackGradient"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   59
         Top             =   360
         Width           =   945
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3360
         Picture         =   "frmTest.frx":2F79E
         Top             =   3720
         Width           =   615
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000010&
         Height          =   4575
         Left            =   120
         Top             =   120
         Width           =   5535
      End
   End
   Begin Test.McToolBar Tabs 
      Height          =   480
      Left            =   240
      TabIndex        =   49
      Top             =   1680
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   4
      ButtonsWidth    =   77
      ButtonsPerRow   =   11
      HoverColor      =   8388608
      TooTipStyle     =   0
      ButtonsMode     =   6
      ButtonsBackColor=   -2147483644
      ButtonsGradientCol=   14737632
      ButtonsGradient =   3
      ButtonCaption1  =   "How to ??"
      ButtonToolTipIcon1=   1
      ButtonCaption2  =   "Properties"
      ButtonToolTipIcon2=   1
      ButtonCaption3  =   "Appearance"
      ButtonToolTipIcon3=   1
      ButtonCaption4  =   "Styles"
      ButtonToolTipIcon4=   1
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmbBorder_Click()
Dim X1 As Long

    For X1 = 0 To 6
    With McToolBar(X1)
        .BorderStyle = cmbBorder.ListIndex
    End With
    Next X1
    
End Sub

Private Sub cmbBtnGrd_Click()
Dim X1 As Long

    For X1 = 0 To 6
    With McToolBar(X1)
        .ButtonsGradient = cmbBtnGrd.ListIndex
    End With
    Next X1
    
End Sub

Private Sub cmbGradient_Click()
Dim X1 As Long

    For X1 = 0 To 6
    With McToolBar(X1)
        .BackGradient = cmbGradient.ListIndex
    End With
    Next X1
    
End Sub


Private Sub picCol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X1 As Long

    For X1 = 0 To 6
    With McToolBar(X1)
    
        If optBack Then .BackColor = picCol.Point(X, Y)
        If optBackGrd Then .BackGradientCol = picCol.Point(X, Y)
        If optBtnBack Then .ButtonsBackColor = picCol.Point(X, Y)
        If optBtnGrad Then .ButtonsGradientCol = picCol.Point(X, Y)
        If optFore Then .ForeColor = picCol.Point(X, Y)
        If optHover Then .HoverColor = picCol.Point(X, Y)
        If optTipBack Then .ToolTipBackCol = picCol.Point(X, Y)
        If optTipFore Then .ToolTipForeCol = picCol.Point(X, Y)
    
    End With
    Next X1
    
End Sub

Private Sub Tabs_Click(ByVal ButtonIndex As Long)
 
 Dim X1  As Long
 
    For X1 = 1 To 4
    With McToolBar(X1)
        Tabs.SetButtonValue X1, BTN_Pressed, False
    End With
    Next X1
    
    picFrame(ButtonIndex - 1).ZOrder (0)
    Tabs.SetButtonValue ButtonIndex, BTN_Pressed, True

End Sub

Private Sub tbApply_Click(ByVal ButtonIndex As Long)
Dim X As Long

    For X = 0 To 6
    With McToolBar(X)
    
        .ButtonsWidth = Val(txtWidth)
        .ButtonsHeight = Val(txtHeight)
        .Enabled = chkEnablectl
        .ButtonsPerRow = Val(txtRow)
        .ButtonsPerRow_Chev = Val(txtRowChev)
        .ShowChevron = chkChev
        .ShowSeperator = chkSep
        .Button_Type = cmbType.ListIndex
        .Button_Index = Val(txtindex)
        .ButtonCaption = txtCaption
        .ToolTipText = txtTooltip
        .ButtonToolTipIcon = cmbIcon.ListIndex
        .TooTipStyle = cmbStyle.ListIndex
        .ButtonEnabled = chkEnabled
        .ButtonPressed = chkPress
        .ButtonIconAllignment = cmbCapAln.ListIndex
    
    End With
    Next X
    
End Sub

Private Sub tbTile_Click(ByVal ButtonIndex As Long)
Dim X1 As Long

    For X1 = 0 To 6
    With McToolBar(X1)
        Set .BackGround = Image2.Picture
    End With
    Next X1
    
End Sub

