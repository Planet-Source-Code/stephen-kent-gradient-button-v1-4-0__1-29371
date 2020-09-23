VERSION 5.00
Object = "*\AGradButton.vbp"
Begin VB.Form frmGradButtonExample 
   Caption         =   "Gradient Button Example"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin GradButton.GradientButton gbAbout 
      Height          =   405
      Left            =   11010
      TabIndex        =   87
      ToolTipText     =   "Click here to display the about information for the control."
      Top             =   4980
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   714
      Appearance      =   0
      Caption         =   "&About"
      CaptionStyle    =   3
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverForeColor  =   16711680
   End
   Begin VB.Frame frElliptical 
      Caption         =   "Elliptical Gradients"
      Height          =   1545
      Left            =   8400
      TabIndex        =   84
      Top             =   3540
      Width           =   2385
      Begin GradButton.GradientButton gbElliptical 
         Height          =   615
         Index           =   0
         Left            =   60
         TabIndex        =   59
         ToolTipText     =   "Ellipitical Gradient Example 1"
         Top             =   240
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   1085
         Caption         =   "Elliptical Example 1"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   65535
         GradientColor2  =   255
         GradientRepetitions=   3.5
         GradientType    =   1
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin GradButton.GradientButton gbElliptical 
         Height          =   615
         Index           =   1
         Left            =   60
         TabIndex        =   60
         ToolTipText     =   "Ellipitical Gradient Example 2"
         Top             =   870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1085
         Appearance      =   0
         Caption         =   "Elliptical Example 2"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientBlendMode=   1
         GradientColor1  =   16777088
         GradientColor2  =   12582912
         GradientRepetitions=   3
         GradientType    =   1
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin GradButton.GradientButton gbElliptical 
         Height          =   615
         Index           =   2
         Left            =   1200
         TabIndex        =   61
         ToolTipText     =   "Ellipitical Gradient Example 3"
         Top             =   870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Appearance      =   2
         Caption         =   "Elliptical Example 3"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   128
         GradientColor2  =   12632319
         GradientRepetitions=   2
         GradientType    =   1
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
   End
   Begin VB.Frame frAutoSize 
      Caption         =   "AutoSizing Background"
      Height          =   1545
      Left            =   8400
      TabIndex        =   83
      Top             =   1950
      Width           =   2385
      Begin GradButton.GradientButton gbAutoSize 
         Height          =   525
         Index           =   0
         Left            =   60
         TabIndex        =   57
         ToolTipText     =   "This resizes to match the background picture size."
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         AutoSize        =   2
         BackPicture     =   "frmGradButtonExample.frx":0000
         Caption         =   "AutoSizing Control to Pic."
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   4
      End
      Begin GradButton.GradientButton gbAutoSize 
         Height          =   615
         Index           =   1
         Left            =   60
         TabIndex        =   58
         ToolTipText     =   "This control resizes the background picture to match the control size."
         Top             =   870
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   1085
         Appearance      =   0
         AutoSize        =   1
         BackPicture     =   "frmGradButtonExample.frx":22C6
         Caption         =   "AutoSizing Picture to Control"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   4
      End
   End
   Begin VB.Frame frOptionButton 
      Caption         =   "Option Button"
      Height          =   2445
      Left            =   10830
      TabIndex        =   85
      Top             =   2490
      Width           =   2625
      Begin GradButton.GradientButton gbOptionButton 
         Height          =   525
         Index           =   0
         Left            =   60
         TabIndex        =   62
         ToolTipText     =   "Option Button Example 1 (All Option Buttons in this example  belong to same group)"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         ButtonType      =   2
         Caption         =   "Option Button 1 Value = True"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   75
         GradientColor1  =   16777215
         GradientColor2  =   4210752
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Value           =   -1  'True
      End
      Begin GradButton.GradientButton gbOptionButton 
         Height          =   525
         Index           =   1
         Left            =   60
         TabIndex        =   63
         ToolTipText     =   "Option Button Example 2 (All Option Buttons in this example  belong to same group)"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         BorderColor     =   8421631
         ButtonType      =   2
         Caption         =   "Option Button 2 Value = False"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   120
         GradientColor1  =   128
         GradientColor2  =   8421631
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin GradButton.GradientButton gbOptionButton 
         Height          =   525
         Index           =   2
         Left            =   60
         TabIndex        =   64
         ToolTipText     =   "Option Button Example 3 (All Option Buttons in this example  belong to same group)"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   2
         BorderColor     =   33023
         ButtonType      =   2
         Caption         =   "Option Button 3 Value = False"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   93.6
         GradientColor1  =   255
         GradientColor2  =   65535
         GradientRepetitions=   2
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin GradButton.GradientButton gbOptionButton 
         Height          =   525
         Index           =   3
         Left            =   60
         TabIndex        =   65
         ToolTipText     =   "Option Button Example 4 (All Option Buttons in this example  belong to same group)"
         Top             =   1860
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   3
         ButtonType      =   2
         Caption         =   "Option Button 4 Value = False"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   169.934
         GradientBlendMode=   1
         GradientColor1  =   16777088
         GradientColor2  =   8388608
         GradientRepetitions=   3
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin GradButton.GradientButton gbOptionButton 
         Height          =   525
         Index           =   4
         Left            =   1320
         TabIndex        =   66
         ToolTipText     =   "Option Button Example 5 (All Option Buttons in this example  belong to same group)"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         BackPicture     =   "frmGradButtonExample.frx":458C
         BorderColor     =   16761087
         ButtonType      =   2
         Caption         =   "Option Button 5 Value = False"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   297
         GradientColor1  =   16744703
         GradientColor2  =   16744576
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   4
      End
      Begin GradButton.GradientButton gbOptionButton 
         Height          =   525
         Index           =   5
         Left            =   1320
         TabIndex        =   67
         ToolTipText     =   "Option Button Example 6 (All Option Buttons in this example  belong to same group)"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         BackPicture     =   "frmGradButtonExample.frx":6852
         BorderColor     =   12648447
         ButtonType      =   2
         Caption         =   "Option Button 6 Value = False"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   65
         GradientColor1  =   49152
         GradientColor2  =   16777088
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   4
      End
      Begin GradButton.GradientButton gbOptionButton 
         Height          =   525
         Index           =   6
         Left            =   1320
         TabIndex        =   68
         ToolTipText     =   "Option Button Example 7 (All Option Buttons in this example  belong to same group)"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   2
         BackPicture     =   "frmGradButtonExample.frx":8B18
         BorderColor     =   8421631
         ButtonType      =   2
         Caption         =   "Option Button 7 Value = False"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   115
         GradientColor1  =   12632319
         GradientColor2  =   8388736
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   4
      End
      Begin GradButton.GradientButton gbOptionButton 
         Height          =   525
         Index           =   7
         Left            =   1320
         TabIndex        =   69
         ToolTipText     =   "Option Button Example 8 (All Option Buttons in this example  belong to same group)"
         Top             =   1860
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   3
         BackPicture     =   "frmGradButtonExample.frx":ADDE
         BorderColor     =   8421631
         ButtonType      =   2
         Caption         =   "Option Button 8 Value = False"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   115
         GradientColor1  =   12632319
         GradientColor2  =   8388736
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   4
      End
   End
   Begin VB.Frame frStateButton 
      Caption         =   "State Button"
      Height          =   1905
      Left            =   9420
      TabIndex        =   74
      Top             =   0
      Width           =   1365
      Begin GradButton.GradientButton gbStateButton 
         Height          =   525
         Index           =   0
         Left            =   60
         TabIndex        =   22
         ToolTipText     =   "State Button Example 1"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         BorderColor     =   8421631
         ButtonType      =   1
         Caption         =   "State Button 1 Value = False"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   235
         GradientColor1  =   128
         GradientColor2  =   8421631
         GradientRepetitions=   2
         GradientType    =   2
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin GradButton.GradientButton gbStateButton 
         Height          =   525
         Index           =   1
         Left            =   60
         TabIndex        =   23
         ToolTipText     =   "State Button Example 2"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         BorderColor     =   33023
         ButtonType      =   1
         Caption         =   "State Button 2 Value = True"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   235
         GradientColor1  =   255
         GradientColor2  =   65535
         GradientRepetitions=   4
         GradientType    =   2
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Value           =   -1  'True
      End
      Begin GradButton.GradientButton gbStateButton 
         Height          =   525
         Index           =   2
         Left            =   60
         TabIndex        =   24
         ToolTipText     =   "State Button Example 3"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   2
         BorderColor     =   16777088
         ButtonType      =   1
         Caption         =   "State Button 3 Value = False"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientAngle   =   235
         GradientBlendMode=   1
         GradientColor1  =   16777088
         GradientColor2  =   8388608
         GradientType    =   1
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
   End
   Begin VB.Frame frCaptionStyle 
      Caption         =   "Caption Style"
      Height          =   1905
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   2625
      Begin GradButton.GradientButton gbCaptionStyle 
         Height          =   525
         Index           =   0
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Standard Caption Style"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Caption         =   "Standard"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16711680
      End
      Begin GradButton.GradientButton gbCaptionStyle 
         Height          =   525
         Index           =   1
         Left            =   60
         TabIndex        =   2
         ToolTipText     =   "Light Inset Caption Style"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         Caption         =   "Inset Light"
         CaptionStyle    =   1
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16711680
      End
      Begin GradButton.GradientButton gbCaptionStyle 
         Height          =   525
         Index           =   2
         Left            =   60
         TabIndex        =   3
         ToolTipText     =   "Heavy Inset Caption Style"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   2
         Caption         =   "Inset Heavy"
         CaptionStyle    =   2
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16711680
      End
      Begin GradButton.GradientButton gbCaptionStyle 
         Height          =   525
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         ToolTipText     =   "Light Raised Caption Style"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Caption         =   "Raised Light"
         CaptionStyle    =   3
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16711680
      End
      Begin GradButton.GradientButton gbCaptionStyle 
         Height          =   525
         Index           =   4
         Left            =   1320
         TabIndex        =   5
         ToolTipText     =   "Heavy Raised Caption Style"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         Caption         =   "Raised Heavy"
         CaptionStyle    =   4
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16711680
      End
      Begin GradButton.GradientButton gbCaptionStyle 
         Height          =   525
         Index           =   5
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   "Drop Shadow Caption Style"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   2
         Caption         =   "Drop Shadow"
         CaptionStyle    =   5
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16711680
      End
   End
   Begin GradButton.GradientButton gbExit 
      Height          =   405
      Left            =   12270
      TabIndex        =   0
      ToolTipText     =   "Exit this program"
      Top             =   4980
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   714
      Appearance      =   0
      Caption         =   "E&xit"
      CaptionStyle    =   3
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverForeColor  =   16711680
   End
   Begin VB.Frame frStyle 
      Caption         =   "Styles"
      Height          =   2715
      Left            =   0
      TabIndex        =   76
      Top             =   1950
      Width           =   8355
      Begin VB.Frame frGraphicalPicture 
         Caption         =   "Graphical Picture"
         Height          =   2445
         Left            =   6930
         TabIndex        =   82
         Top             =   210
         Width           =   1365
         Begin GradButton.GradientButton gbGraphicalPicture 
            Height          =   525
            Index           =   0
            Left            =   60
            TabIndex        =   53
            ToolTipText     =   "Graphical Picture Display Style Example 1"
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            BackPicture     =   "frmGradButtonExample.frx":D0A4
            Caption         =   "Graphical Gradient Ex. 1"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   75
            GradientColor1  =   16777215
            GradientColor2  =   4210752
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   5
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton gbGraphicalPicture 
            Height          =   525
            Index           =   1
            Left            =   60
            TabIndex        =   54
            ToolTipText     =   "Graphical Picture Display Style Example 2"
            Top             =   780
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Appearance      =   0
            BackPicture     =   "frmGradButtonExample.frx":F36A
            BorderColor     =   8421631
            Caption         =   "Graphical Gradient Ex. 2"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DisabledPicture =   "frmGradButtonExample.frx":11630
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownPicture     =   "frmGradButtonExample.frx":1230A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   120
            GradientColor1  =   128
            GradientColor2  =   8421631
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverPicture    =   "frmGradButtonExample.frx":12FE4
            Picture         =   "frmGradButtonExample.frx":13CBE
            Style           =   5
         End
         Begin GradButton.GradientButton gbGraphicalPicture 
            Height          =   525
            Index           =   2
            Left            =   60
            TabIndex        =   55
            ToolTipText     =   "Graphical Picture Display Style Example 3"
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Appearance      =   2
            BackPicture     =   "frmGradButtonExample.frx":14998
            BorderColor     =   33023
            Caption         =   "Graphical Gradient Ex. 3"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownPicture     =   "frmGradButtonExample.frx":16C5E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   93.6
            GradientColor1  =   255
            GradientColor2  =   65535
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmGradButtonExample.frx":17938
            Style           =   5
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton gbGraphicalPicture 
            Height          =   525
            Index           =   3
            Left            =   60
            TabIndex        =   56
            ToolTipText     =   "Graphical Picture Display Style Example 4"
            Top             =   1860
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Appearance      =   3
            BackPicture     =   "frmGradButtonExample.frx":18612
            BorderColor     =   16777088
            Caption         =   "Graphical Gradient Ex. 4"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   169.934
            GradientColor1  =   16777088
            GradientColor2  =   8388608
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverPicture    =   "frmGradButtonExample.frx":1A8D8
            Picture         =   "frmGradButtonExample.frx":1B5B2
            Style           =   5
         End
      End
      Begin VB.Frame frPicture 
         Caption         =   "Picture"
         Height          =   2445
         Left            =   5550
         TabIndex        =   81
         Top             =   210
         Width           =   1365
         Begin GradButton.GradientButton gbPicture 
            Height          =   525
            Index           =   0
            Left            =   60
            TabIndex        =   49
            ToolTipText     =   "Picture Display Style Example 1"
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            BackPicture     =   "frmGradButtonExample.frx":1C28C
            Caption         =   "Picture Ex. 1"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   75
            GradientColor1  =   16777215
            GradientColor2  =   4210752
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   4
         End
         Begin GradButton.GradientButton gbPicture 
            Height          =   525
            Index           =   1
            Left            =   60
            TabIndex        =   50
            ToolTipText     =   "Picture Display Style Example 2"
            Top             =   780
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Appearance      =   0
            BackPicture     =   "frmGradButtonExample.frx":1E552
            BorderColor     =   8421631
            Caption         =   "Picture Ex. 2"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   120
            GradientColor1  =   128
            GradientColor2  =   8421631
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   4
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton gbPicture 
            Height          =   525
            Index           =   2
            Left            =   60
            TabIndex        =   51
            ToolTipText     =   "Picture Display Style Example 3"
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Appearance      =   2
            BackPicture     =   "frmGradButtonExample.frx":20818
            BorderColor     =   33023
            Caption         =   "Picture Ex. 3"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   93.6
            GradientColor1  =   255
            GradientColor2  =   65535
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   4
         End
         Begin GradButton.GradientButton gbPicture 
            Height          =   525
            Index           =   3
            Left            =   60
            TabIndex        =   52
            ToolTipText     =   "Picture Display Style Example 4"
            Top             =   1860
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Appearance      =   3
            BackPicture     =   "frmGradButtonExample.frx":22ADE
            BorderColor     =   16777088
            Caption         =   "Picture Ex. 4"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   165.934
            GradientColor1  =   16777088
            GradientColor2  =   8388608
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   4
            UseHover        =   0   'False
         End
      End
      Begin VB.Frame frGraphicalGradient 
         Caption         =   "Graphical Gradient"
         Height          =   2445
         Left            =   4170
         TabIndex        =   80
         Top             =   210
         Width           =   1365
         Begin GradButton.GradientButton gbGraphicalGradient 
            Height          =   525
            Index           =   4
            Left            =   60
            TabIndex        =   45
            ToolTipText     =   "Graphical Gradient Display Style Example 1"
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Caption         =   "Graphical Gradient Ex. 1"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   75
            GradientColor1  =   16777215
            GradientColor2  =   4210752
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   3
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton gbGraphicalGradient 
            Height          =   525
            Index           =   0
            Left            =   60
            TabIndex        =   46
            ToolTipText     =   "Graphical Gradient Display Style Example 2"
            Top             =   780
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Appearance      =   0
            BorderColor     =   8421631
            Caption         =   "Graphical Gradient Ex. 2"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DisabledPicture =   "frmGradButtonExample.frx":24DA4
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownPicture     =   "frmGradButtonExample.frx":25A7E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   120
            GradientColor1  =   128
            GradientColor2  =   8421631
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverPicture    =   "frmGradButtonExample.frx":26758
            Picture         =   "frmGradButtonExample.frx":27432
            Style           =   3
         End
         Begin GradButton.GradientButton gbGraphicalGradient 
            Height          =   525
            Index           =   1
            Left            =   60
            TabIndex        =   47
            ToolTipText     =   "Graphical Gradient Display Style Example 3"
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Appearance      =   2
            BorderColor     =   33023
            Caption         =   "Graphical Gradient Ex. 3"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownPicture     =   "frmGradButtonExample.frx":2810C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   93.6
            GradientColor1  =   255
            GradientColor2  =   65535
            GradientRepetitions=   2
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmGradButtonExample.frx":28DE6
            Style           =   3
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton gbGraphicalGradient 
            Height          =   525
            Index           =   2
            Left            =   60
            TabIndex        =   48
            ToolTipText     =   "Graphical Gradient Display Style Example 4"
            Top             =   1860
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Appearance      =   3
            BorderColor     =   16777088
            Caption         =   "Graphical Gradient Ex. 4"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   169.934
            GradientBlendMode=   1
            GradientColor1  =   16777088
            GradientColor2  =   8388608
            GradientRepetitions=   3
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverPicture    =   "frmGradButtonExample.frx":29AC0
            Picture         =   "frmGradButtonExample.frx":2A79A
            Style           =   3
         End
      End
      Begin VB.Frame frGradient 
         Caption         =   "Gradient"
         Height          =   2445
         Left            =   2790
         TabIndex        =   79
         Top             =   210
         Width           =   1365
         Begin GradButton.GradientButton gbGradient 
            Height          =   525
            Index           =   0
            Left            =   60
            TabIndex        =   41
            ToolTipText     =   "Gradient Display Style Example 1"
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Caption         =   "Gradient Ex. 1"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   75
            GradientColor1  =   16777215
            GradientColor2  =   4210752
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
         End
         Begin GradButton.GradientButton gbGradient 
            Height          =   525
            Index           =   1
            Left            =   60
            TabIndex        =   42
            ToolTipText     =   "Gradient Display Style Example 2"
            Top             =   780
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Appearance      =   0
            BorderColor     =   8421631
            Caption         =   "Gradient Ex. 2"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   120
            GradientColor1  =   128
            GradientColor2  =   8421631
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton gbGradient 
            Height          =   525
            Index           =   2
            Left            =   60
            TabIndex        =   43
            ToolTipText     =   "Gradient Display Style Example 3"
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Appearance      =   2
            BorderColor     =   33023
            Caption         =   "Gradient Ex. 3"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   93.6
            GradientColor1  =   255
            GradientColor2  =   65535
            GradientRepetitions=   2
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
         End
         Begin GradButton.GradientButton gbGradient 
            Height          =   525
            Index           =   3
            Left            =   60
            TabIndex        =   44
            ToolTipText     =   "Gradient Display Style Example 4"
            Top             =   1860
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Appearance      =   3
            BorderColor     =   16777088
            Caption         =   "Gradient Ex. 4"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientAngle   =   165.934
            GradientBlendMode=   1
            GradientColor1  =   16777088
            GradientColor2  =   8388608
            GradientRepetitions=   3
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
            UseHover        =   0   'False
         End
      End
      Begin VB.Frame frStyleGraphical 
         Caption         =   "Graphical"
         Height          =   2445
         Left            =   1410
         TabIndex        =   78
         Top             =   210
         Width           =   1365
         Begin GradButton.GradientButton gbGraphical 
            Height          =   525
            Index           =   0
            Left            =   60
            TabIndex        =   37
            ToolTipText     =   "Graphical Display Style Example 1"
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Caption         =   "Graphical Ex. 1"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton gbGraphical 
            Height          =   525
            Index           =   1
            Left            =   60
            TabIndex        =   38
            ToolTipText     =   "Graphical Display Style Example 2"
            Top             =   780
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Appearance      =   0
            BackColor       =   255
            Caption         =   "Graphical Ex. 2"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DisabledPicture =   "frmGradButtonExample.frx":2B474
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownPicture     =   "frmGradButtonExample.frx":2C14E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverPicture    =   "frmGradButtonExample.frx":2CE28
            Picture         =   "frmGradButtonExample.frx":2DB02
            Style           =   1
         End
         Begin GradButton.GradientButton gbGraphical 
            Height          =   525
            Index           =   2
            Left            =   60
            TabIndex        =   39
            ToolTipText     =   "Graphical Display Style Example 3"
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Appearance      =   2
            BackColor       =   16576
            Caption         =   "Graphical Ex. 3"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DownPicture     =   "frmGradButtonExample.frx":2E7DC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmGradButtonExample.frx":2F4B6
            Style           =   1
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton gbGraphical 
            Height          =   525
            Index           =   3
            Left            =   60
            TabIndex        =   40
            ToolTipText     =   "Graphical Display Style Example 4"
            Top             =   1860
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Alignment       =   8
            Appearance      =   3
            BackColor       =   16711680
            BevelIntensity  =   40
            Caption         =   "Graphical Ex. 4"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverPicture    =   "frmGradButtonExample.frx":30190
            Picture         =   "frmGradButtonExample.frx":30E6A
            Style           =   1
         End
      End
      Begin VB.Frame frStyleStandard 
         Caption         =   "Standard"
         Height          =   2445
         Left            =   30
         TabIndex        =   77
         Top             =   210
         Width           =   1365
         Begin GradButton.GradientButton gbStandard 
            Height          =   525
            Index           =   0
            Left            =   60
            TabIndex        =   33
            ToolTipText     =   "Standard Display Style Example 1"
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Caption         =   "Standard Ex. 1"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin GradButton.GradientButton gbStandard 
            Height          =   525
            Index           =   1
            Left            =   60
            TabIndex        =   34
            ToolTipText     =   "Standard Display Style Example 2"
            Top             =   780
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Appearance      =   0
            BackColor       =   255
            Caption         =   "Standard Ex. 2"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseHover        =   0   'False
         End
         Begin GradButton.GradientButton gbStandard 
            Height          =   525
            Index           =   2
            Left            =   60
            TabIndex        =   35
            ToolTipText     =   "Standard Display Style Example 3"
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Appearance      =   2
            BackColor       =   16576
            Caption         =   "Standard Ex. 3"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin GradButton.GradientButton gbStandard 
            Height          =   525
            Index           =   3
            Left            =   60
            TabIndex        =   36
            ToolTipText     =   "Standard Display Style Example 4"
            Top             =   1860
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   926
            Appearance      =   3
            BackColor       =   16711680
            BevelIntensity  =   40
            Caption         =   "Standard Ex. 4"
            BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseHover        =   0   'False
         End
      End
   End
   Begin VB.Frame frFont 
      Caption         =   "Font"
      Height          =   1905
      Left            =   8010
      TabIndex        =   73
      Top             =   0
      Width           =   1365
      Begin GradButton.GradientButton gbFont 
         Height          =   525
         Index           =   0
         Left            =   60
         TabIndex        =   19
         ToolTipText     =   "All Standard Fonts"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Caption         =   "Standard"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledFontEnabled=   -1  'True
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DownFontEnabled =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverFontEnabled=   -1  'True
      End
      Begin GradButton.GradientButton gbFont 
         Height          =   525
         Index           =   1
         Left            =   60
         TabIndex        =   20
         ToolTipText     =   "MS SanSerif Font, Verdana Hover font, Times New Roman Down font, and Courier New Disabled font."
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         Caption         =   "Example 1"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         DisabledFontEnabled=   -1  'True
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         DownFontEnabled =   -1  'True
         DownForeColor   =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         HoverFontEnabled=   -1  'True
         HoverForeColor  =   16711680
      End
      Begin GradButton.GradientButton gbFont 
         Height          =   525
         Index           =   2
         Left            =   60
         TabIndex        =   21
         ToolTipText     =   "Arial Font, Courier New Hover font, System Down font, and MS Sans Serif Disabled font."
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   2
         Caption         =   "Example 2"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledFontEnabled=   -1  'True
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   -1  'True
         EndProperty
         DownFontEnabled =   -1  'True
         DownForeColor   =   49344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   12583104
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverFontEnabled=   -1  'True
         HoverForeColor  =   16776960
      End
   End
   Begin VB.Frame frColor 
      Caption         =   "Color"
      Height          =   1905
      Left            =   6600
      TabIndex        =   72
      Top             =   0
      Width           =   1365
      Begin GradButton.GradientButton gbColor 
         Height          =   525
         Index           =   0
         Left            =   60
         TabIndex        =   16
         ToolTipText     =   "Standard Text Colors"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Caption         =   "Standard"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GradButton.GradientButton gbColor 
         Height          =   525
         Index           =   1
         Left            =   60
         TabIndex        =   17
         ToolTipText     =   "Red forecolor, Blue Hover Color, Green Down color, and System Disabled Text Color when disabled."
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         Caption         =   "Example 1"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DownForeColor   =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16711680
      End
      Begin GradButton.GradientButton gbColor 
         Height          =   525
         Index           =   2
         Left            =   60
         TabIndex        =   18
         ToolTipText     =   "Purple forecolor, Cyan Hover Color, Dirty Yellow Down color, and System Disabled Text Color when disabled."
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   2
         Caption         =   "Example 2"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DownForeColor   =   49344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12583104
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   16776960
      End
   End
   Begin VB.Frame frCaptionAlign 
      Caption         =   "Caption Align"
      Height          =   1905
      Left            =   2670
      TabIndex        =   71
      Top             =   0
      Width           =   3885
      Begin GradButton.GradientButton gbCaptionAlign 
         Height          =   525
         Index           =   3
         Left            =   1320
         TabIndex        =   10
         ToolTipText     =   "Text Positioned at Center Top"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Alignment       =   6
         Caption         =   "Center Top"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
      Begin GradButton.GradientButton gbCaptionAlign 
         Height          =   525
         Index           =   0
         Left            =   60
         TabIndex        =   7
         ToolTipText     =   "Text Positioned at Left Top"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Alignment       =   0
         Caption         =   "Left Top"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
      Begin GradButton.GradientButton gbCaptionAlign 
         Height          =   525
         Index           =   1
         Left            =   60
         TabIndex        =   8
         ToolTipText     =   "Text Positioned at Left Middle"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Alignment       =   1
         Appearance      =   0
         Caption         =   "Left Middle"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
      Begin GradButton.GradientButton gbCaptionAlign 
         Height          =   525
         Index           =   2
         Left            =   60
         TabIndex        =   9
         ToolTipText     =   "Text Positioned at Left Bottom"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Alignment       =   2
         Appearance      =   2
         Caption         =   "Left Bottom"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
      Begin GradButton.GradientButton gbCaptionAlign 
         Height          =   525
         Index           =   4
         Left            =   1320
         TabIndex        =   11
         ToolTipText     =   "Text Positioned at Center Middle"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         Caption         =   "Center Midde"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
      Begin GradButton.GradientButton gbCaptionAlign 
         Height          =   525
         Index           =   5
         Left            =   1320
         TabIndex        =   12
         ToolTipText     =   "Text Positioned at Center Bottom"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Alignment       =   8
         Appearance      =   2
         Caption         =   "Center Bottom"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
      Begin GradButton.GradientButton gbCaptionAlign 
         Height          =   525
         Index           =   6
         Left            =   2580
         TabIndex        =   13
         ToolTipText     =   "Text Positioned at Right Top"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Alignment       =   3
         Caption         =   "Right Top"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
      Begin GradButton.GradientButton gbCaptionAlign 
         Height          =   525
         Index           =   7
         Left            =   2580
         TabIndex        =   14
         ToolTipText     =   "Text Positioned at Right Middle"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Alignment       =   4
         Appearance      =   0
         Caption         =   "Right Middle"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
      Begin GradButton.GradientButton gbCaptionAlign 
         Height          =   525
         Index           =   8
         Left            =   2580
         TabIndex        =   15
         ToolTipText     =   "Text Positioned at Right Bottom"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Alignment       =   5
         Appearance      =   2
         Caption         =   "Right Bottom"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
   End
   Begin VB.Frame frAppearance 
      Caption         =   "Appearances"
      Height          =   2445
      Left            =   10830
      TabIndex        =   75
      Top             =   0
      Width           =   2625
      Begin GradButton.GradientButton gbAppearance 
         Height          =   525
         Index           =   0
         Left            =   60
         TabIndex        =   25
         ToolTipText     =   "3D Border style with Hover Display"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Caption         =   "3-D with Hover"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
      End
      Begin GradButton.GradientButton gbAppearance 
         Height          =   525
         Index           =   1
         Left            =   60
         TabIndex        =   27
         ToolTipText     =   "Flat Border style with Hover Display"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         Caption         =   "Flat with Hover"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483634
      End
      Begin GradButton.GradientButton gbAppearance 
         Height          =   525
         Index           =   2
         Left            =   60
         TabIndex        =   29
         ToolTipText     =   "Etched Border style with Hover Display"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   2
         Caption         =   "Etched with Hover"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483631
      End
      Begin GradButton.GradientButton gbAppearance 
         Height          =   525
         Index           =   3
         Left            =   1320
         TabIndex        =   26
         ToolTipText     =   "3D Border style without Hover Display"
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Caption         =   "3-D w/o Hover"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483635
         UseHover        =   0   'False
      End
      Begin GradButton.GradientButton gbAppearance 
         Height          =   525
         Index           =   4
         Left            =   1320
         TabIndex        =   28
         ToolTipText     =   "Flat Border style without Hover Display"
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   0
         Caption         =   "Flat w/o Hover"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483634
         UseHover        =   0   'False
      End
      Begin GradButton.GradientButton gbAppearance 
         Height          =   525
         Index           =   5
         Left            =   1320
         TabIndex        =   30
         ToolTipText     =   "Etched Border style without Hover Display"
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   2
         Caption         =   "Etched w/o Hover"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483631
         UseHover        =   0   'False
      End
      Begin GradButton.GradientButton gbAppearance 
         Height          =   525
         Index           =   6
         Left            =   60
         TabIndex        =   31
         ToolTipText     =   "Beveled Border style with Hover Display"
         Top             =   1860
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   3
         Caption         =   "Bevel with Hover"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483624
      End
      Begin GradButton.GradientButton gbAppearance 
         Height          =   525
         Index           =   7
         Left            =   1320
         TabIndex        =   32
         ToolTipText     =   "Beveled Border style without Hover Display"
         Top             =   1860
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   926
         Appearance      =   3
         Caption         =   "Bevel w/o Hover"
         BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverForeColor  =   -2147483624
         UseHover        =   0   'False
      End
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmGradButtonExample.frx":31B44
      Height          =   615
      Left            =   0
      TabIndex        =   86
      Top             =   4680
      Width           =   8325
   End
End
Attribute VB_Name = "frmGradButtonExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    gbAutoSize(0).Height = gbAutoSize(1).Height
    gbAutoSize(0).Width = gbAutoSize(1).Width
End Sub

Private Sub gbAbout_Click()
    gbAbout.About
End Sub

Private Sub gbExit_Click()
    Unload Me
End Sub

Private Sub gbOptionButton_ValueChanged(Index As Integer, New_Value As Boolean)
    gbOptionButton(Index).Caption = "Option Button " & CStr(Index + 1) & " Value = " & New_Value
End Sub

Private Sub gbStateButton_ValueChanged(Index As Integer, New_Value As Boolean)
    gbStateButton(Index).Caption = "State Button " & CStr(Index + 1) & " Value = " & New_Value
End Sub
