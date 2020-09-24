VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Ppal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert - Extract WORD files into/from an Access DataBase"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDialogo 
      Left            =   5070
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton B_Examinar 
      BackColor       =   &H000000FF&
      Caption         =   "..."
      Height          =   345
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FEFBED&
      Caption         =   "&Extract all documents from DB"
      Height          =   525
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   390
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Save Doc. into DataBase"
      Default         =   -1  'True
      Height          =   525
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   390
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6225
   End
End
Attribute VB_Name = "frm_Ppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub B_Examinar_Click()
On Error GoTo Controlar
    CDialogo.CancelError = True
    CDialogo.ShowOpen
    Text1.Text = CDialogo.FileName
Controlar:
    If Err.Number = 0 Or Err.Number = 37 Then Exit Sub
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub Command1_Click()
If SaveFileToDB(Text1.Text, "Fichero", "Lista", App.Path & "\documentos.mdb") Then MsgBox "Stored into DB success!", vbInformation, "Store Operation"
End Sub

Private Sub Command2_Click()
    If LoadFileFromDB(App.Path & "\temp", "Fichero", "Lista", App.Path & "\documentos.mdb") Then MsgBox "All .doc's extracted successfully!!", vbInformation, "Extract Operation"
End Sub
