VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   ScaleHeight     =   7065
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "nuevo"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "eliminar"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "nuevo"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   855
      Left            =   2520
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "edad"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "ejercicio"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "nutricion"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "dieta"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "estudiante"
      Top             =   0
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante\Desktop\l\basededatos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "es"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "carne"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "apellido"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label musica 
      Caption         =   "nombre"
      DataField       =   "musica"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Add.save
End Sub

Private Sub Command2_Click()
Add.delitet

End Sub

Private Sub Command3_Click()
Add.new

End Sub
