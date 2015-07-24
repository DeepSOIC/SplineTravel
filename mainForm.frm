VERSION 5.00
Begin VB.Form mainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SplineTravel"
   ClientHeight    =   4320
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   9490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcessFile 
      Caption         =   "Go"
      CausesValidation=   0   'False
      Height          =   840
      Left            =   5210
      TabIndex        =   13
      Top             =   3290
      Width           =   1890
   End
   Begin VB.TextBox Text3 
      Height          =   400
      Left            =   1440
      TabIndex        =   11
      Text            =   "200"
      Top             =   1520
      Width           =   1260
   End
   Begin VB.TextBox Text2 
      Height          =   320
      Left            =   1410
      TabIndex        =   8
      Text            =   "5"
      Top             =   2670
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   400
      Left            =   1420
      TabIndex        =   5
      Text            =   "2000"
      Top             =   2140
      Width           =   1260
   End
   Begin VB.TextBox txtFNOut 
      Height          =   410
      Left            =   1200
      TabIndex        =   3
      Text            =   "txtFNOut"
      Top             =   480
      Width           =   5590
   End
   Begin VB.TextBox txtFNIn 
      Height          =   360
      Left            =   1190
      TabIndex        =   0
      Text            =   "txtFNIn"
      Top             =   50
      Width           =   5610
   End
   Begin VB.Label Label8 
      Caption         =   "mm/s2"
      Height          =   240
      Left            =   2830
      TabIndex        =   12
      Top             =   1550
      Width           =   920
   End
   Begin VB.Label Label7 
      Caption         =   "speed limit"
      Height          =   340
      Left            =   190
      TabIndex        =   10
      Top             =   1570
      Width           =   1210
   End
   Begin VB.Label Label6 
      Caption         =   "mm/s"
      Height          =   240
      Left            =   2930
      TabIndex        =   9
      Top             =   2690
      Width           =   760
   End
   Begin VB.Label Label5 
      Caption         =   "jerk"
      Height          =   280
      Left            =   150
      TabIndex        =   7
      Top             =   2670
      Width           =   1070
   End
   Begin VB.Label Label4 
      Caption         =   "mm/s2"
      Height          =   240
      Left            =   2810
      TabIndex        =   6
      Top             =   2170
      Width           =   920
   End
   Begin VB.Label Label3 
      Caption         =   "accelleration"
      Height          =   340
      Left            =   170
      TabIndex        =   4
      Top             =   2190
      Width           =   1210
   End
   Begin VB.Label Label2 
      Caption         =   "output"
      Height          =   380
      Left            =   50
      TabIndex        =   2
      Top             =   480
      Width           =   1090
   End
   Begin VB.Label Label1 
      Caption         =   "input"
      Height          =   270
      Left            =   40
      TabIndex        =   1
      Top             =   40
      Width           =   1020
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

