VERSION 2.00
Begin Form About 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   ClientHeight    =   3330
   ClientLeft      =   1950
   ClientTop       =   1800
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Height          =   3705
   Icon            =   ABOUT.FRX:0000
   Left            =   1905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4560
   Top             =   1470
   Width           =   4650
   Begin CommandButton BtnOk 
      Caption         =   "&Ok!"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin SSFrame FrameCopyright 
      Alignment       =   2  'Center
      Caption         =   "Copyright"
      Font3D          =   3  'Inset w/light shading
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
      Begin Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This program is PUBLIC DOMAIN freeware. You may use and copy this program freely. You may NOT make profit on its distribution. Please see the WSANET.HLP file for the full copyright agreement."
         FontBold        =   0   'False
         FontItalic      =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
   End
   Begin SSPanel PanelAppTitle 
      Alignment       =   6  'Center - TOP
      BackColor       =   &H00C0C0C0&
      Caption         =   "PanelAppTitle"
      Font3D          =   1  'Raised w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   480
         Left            =   120
         Picture         =   ABOUT.FRX:0302
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   120
         Width           =   480
      End
      Begin SSPanel Panel3D2 
         Alignment       =   6  'Center - TOP
         BackColor       =   &H00C0C0C0&
         BevelOuter      =   1  'Inset
         Caption         =   "By: Ian C. Blenke"
         Font3D          =   1  'Raised w/light shading
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
   Begin Label LabelAppDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "LabelAppDescription"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
End
Option Explicit

Sub BtnOk_Click ()

    gsMISCAboutUnload

End Sub

