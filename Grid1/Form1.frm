VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStats 
   Caption         =   "Data"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Read"
      Height          =   495
      Left            =   10200
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   5655
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9975
      _Version        =   393216
      Rows            =   50
      Cols            =   6
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Label Label1 
      Caption         =   "New South wales Covid-19 cases"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   8895
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim fileName As String
    Dim filenum As Integer
    fileName = "c:\covidNsw.csv"
    filenum = FreeFile
    
    Dim fileData As String
    Dim fileLines() As String
    Dim fileColumns() As String
    
    Dim i As Integer
    Dim j As Integer
    
    Open fileName For Input As #filenum
    fileData = Input(LOF(filenum), #filenum)
    Close #filenum
    fileLines = Split(fileData, vbLf)
    
   
    For i = 0 To UBound(fileLines) - 1
        'Split each column into an array
        fileColumns = Split(fileLines(i), ",")
        'Loop through each column
        For j = 0 To UBound(fileColumns) - 1
            If j = 0 Then
            Grid.TextMatrix(i, j) = i
            Else
            'Text1.Text = fileColumns(0)
            Grid.TextMatrix(i, j) = fileColumns(j - 1)
            End If
        Next j
    Next i
End Sub

Private Sub Form_Load()
    Grid.Cols = 6
    Grid.Rows = 10000
    Grid.ColWidth(0) = 550
    Grid.ColWidth(1) = 950
    Grid.ColWidth(2) = 820
    Grid.ColWidth(3) = 5300
    Grid.ColWidth(4) = 800
    Grid.ColWidth(5) = 2000
End Sub
