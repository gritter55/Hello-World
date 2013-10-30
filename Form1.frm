VERSION 5.00
Begin VB.Form frmStatus 
   Caption         =   "Status"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblStatus 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "GMT 800 Recap Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
