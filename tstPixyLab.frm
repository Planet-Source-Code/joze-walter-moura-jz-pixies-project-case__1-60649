VERSION 5.00
Begin VB.Form tstPixyLab 
   Caption         =   "JZ: Mouse Events on Different Image Regions"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label20 
      Caption         =   "Mouse"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2565
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Now is"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   3285
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Was"
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   8
      Top             =   2745
      Width           =   330
   End
   Begin VB.Label LbWas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   195
      TabIndex        =   7
      Top             =   2955
      Width           =   1080
   End
   Begin VB.Label Pixy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   4365
      MouseIcon       =   "tstPixyLab.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3540
      Width           =   1260
   End
   Begin VB.Label Pixy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Index           =   3
      Left            =   4875
      MouseIcon       =   "tstPixyLab.frx":0184
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   375
      Width           =   780
   End
   Begin VB.Label Pixy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   2
      Left            =   1545
      MouseIcon       =   "tstPixyLab.frx":0308
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1815
      Width           =   1110
   End
   Begin VB.Label Pixy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   1
      Left            =   4170
      MouseIcon       =   "tstPixyLab.frx":048C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1875
      Width           =   510
   End
   Begin VB.Label Pixy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2550
      MouseIcon       =   "tstPixyLab.frx":0610
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "This is the thumb finger!"
      Top             =   840
      Width           =   510
   End
   Begin VB.Label LbClone 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clone"
      Height          =   360
      Left            =   255
      TabIndex        =   1
      Top             =   120
      Width           =   780
   End
   Begin VB.Label LbNow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   3495
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   3600
      Left            =   885
      Picture         =   "tstPixyLab.frx":0794
      Top             =   255
      Width           =   4800
   End
End
Attribute VB_Name = "tstPixyLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'.-----------.------------------------------------------------------
' Case Study : JZ: Anywhere Mouse Regions
' Edition    : 22-May-2005
' Author     : JOZE Walter de Moura - RIO DE JANEIRO, BRASIL.
'            : me: www.joze.kit.net   or   qualyum@globo.com
'            :
'            : This case is beginners dedicated but would be
'            : appreciated by all professionals due very
'            : simple code solutions.
'            :
' Application: MouseOn, MouseOut, MouseLeave, e+ on different
'            : regions of an any Image Control in simple code:
'            : no APIs, no DLLs, no .Bas - only in Form1
'            : design and easy normal procedures.
'            :
' Sample     : We are showing a photo and we would like do some
'            : procedures according mouse in certain regions of
'            : the image by clicking, pointing, leaving, being
'            : in or out, etc.
'            :
'            : Well, let's going to steps:
'            :
'         01 : A normal Form1; I've named tstPixyLab
'            :
'         02 : Put a Image Control, select the desired Picture.
'            : I've named it "img" and Picture was a joked
'            : "hot dog".
'            :
'         03 : Put a Label Control, anywhere, any size.
'            : I've named it "LbClone", setting BackStyle = 0
'            : (transparent) and Caption = "Clone" (this is the
'            : best for visual design; at Form_Load subroutine,
'            : you will see adjusting procedure:
'            : a) Setting Caption="", BorderStyle=0,
'            :    Appearence=0 (flat); and
'            : b) The label is sized the same as Image.
'            :
'            : Note: This transparent label will be the mouse
'            : controler "Out" the regions, avoiding Image
'            : mouse senses maybe failled at runtime.
'            :
'         04 : Create the 1st "Pixy" label for 1st region:
'            : Put a Label Control inside the image frame,
'            : setting BackStyle=0 (transparent),
'            : Appearence=0 (flat).
'            : I've named it "Pixy", and Caption="0" and
'            : used to delimit a rectangular region - where is
'            : the Thumb pressuring the "hot-dog".
'            : Caption was a optional design setting only for
'            : visual purposes. We will remove borders and
'            : text captions at Form_Load routine, see it.
'            : Ah, another embellishment:
'            : I've designed the "hand.cur" as MouseIcon and
'            : Set MousePointer = 99#.
'            :
'         05 : Now, as often needed, delimiting other regions.
'            : Duplicate the "Pixy" Label one for each region
'            : you wishes control, by copying (ctrl-C) and
'            : pasting (ctrl-V) in different locals.
'            : At first copy, you must aswer "Yes" to
'            : "Control Array...?" message.
'            : I've did it and changed respective Captions to
'            : remember Pixy's indexes.
'            :
'         06 : Other Controls are about my test subjects. Two
'            : variable labels, LbNow and LbWas, mode express
'            : the logical result of mouse position analysis.
'            : Of course, your applications will treat it with
'            : different manners.
'            : I've put it on
'            : 1=Tail 2=Flag Stars 3=Head 4=Photo Date
'            : (0=Thumb, the first).
'            :
'    Logical :
'      Bases : Despite be Transparent, all Mouse events will
'            : be sent through these Labels.
'            : In this code, we get "where IsMouseOver" using
'            : the Pixy_MouseMove(index) subroutine, where
'            : the received Index will indicates the Region.
'            : When LbClone_MouseMove means "MouseLeave"s the
'            : last pointed region and "the MouseIsOut all"!
'            :
'   Possible :
'      to do : Tooltips, Graphic Menus, Games, etc.
'            : I've made a ToolTipText for Region 0=Thumb.
'            :
'    License : Freeware - you may distribute, alter, sold, anything
'            : as you want. This code is for you, don't it?
'            : No credits too - the source is only a tip to
'            : VB community.
'            :
'            : Enjoy,
'            :
'            : Joze.
'            :
'`------------------------------------------------------------------'

Private Sub Form_Load()
   Dim i As Integer
'Clone preparing (would be setting at design time but ...)
   LbClone.BorderStyle = 0 ' no Board
   LbClone.Appearance = 0 'flat
   LbClone.Caption = "" 'no text
'be sure same image size is a fundamental one
   LbClone.Top = Img.Top
   LbClone.Left = Img.Left
   LbClone.Width = Img.Width
   LbClone.Height = Img.Height

'Now, pixies preparing:
'borders and caption was for design visual purposes
   For i = 0 To 4 ' we have 5 regions
     Pixy(i).Caption = "" 'no text
     Pixy(i).BorderStyle = 0 'no border
   Next i
End Sub

'This is a simple model routine to do anything
'when MouseLeave occurs
Private Sub Swapit()
   Const MouseLeave = "Out/Leaves" ' expression for mouse out stat
   If Len(Trim(LbNow.Caption)) <> 0 Then ' there was anything
      If LbNow.Caption <> MouseLeave Then
         LbWas.Caption = LbNow.Caption
         LbNow.Caption = MouseLeave
      End If
   End If
End Sub

'by been back all pixies, this will be
'the same as each Pixy(n).MouseLeave
Private Sub LbClone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Swapit
End Sub

'MouseIsOver any Region: what will be that?
Private Sub Pixy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Index
     Case 0
        LbNow.Caption = "on Thumb"
     Case 1
        LbNow.Caption = "on Tail"
     Case 2
        LbNow.Caption = "on Head"
     Case 3
        LbNow.Caption = "on Stars"
     Case 4
        LbNow.Caption = "on Date"
   End Select
End Sub

'---------------------------------------------------------
'  THAT'S ALL, FOLKS!
'=========================================================
