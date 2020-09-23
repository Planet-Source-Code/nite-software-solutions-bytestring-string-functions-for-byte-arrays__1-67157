VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ByteString"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Byte Concatenation (Append) "
      Height          =   1335
      Left            =   120
      TabIndex        =   36
      Top             =   5640
      Width           =   6615
      Begin VB.TextBox txtConcat 
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Text            =   "Hello"
         Top             =   360
         Width           =   6135
      End
      Begin VB.TextBox txtAppend 
         Height          =   285
         Left            =   960
         TabIndex        =   38
         Text            =   ", world!"
         Top             =   840
         Width           =   4335
      End
      Begin VB.CommandButton cmdOKAppend 
         Caption         =   "OK"
         Height          =   285
         Left            =   5640
         TabIndex        =   37
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Append:"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   40
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame fraIsNumeric 
      Caption         =   " IsNumeric Function "
      Height          =   1815
      Left            =   3480
      TabIndex        =   33
      Top             =   3720
      Width           =   3255
      Begin VB.TextBox txtIsNum 
         Height          =   285
         Left            =   240
         TabIndex        =   35
         Text            =   "1234567890"
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdOKNum 
         Caption         =   "OK"
         Height          =   285
         Left            =   2280
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame fraRev 
      Caption         =   " InStrRev Function "
      Height          =   1815
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   3255
      Begin VB.TextBox txtFindRev 
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Text            =   "hijklm"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtInStrRev 
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Text            =   "AbcdefGhiJklmnoPqrsTuVWXyZ"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtStartRev 
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdOKRev 
         Caption         =   "OK"
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox chkCaseRev 
         Caption         =   "Case sensitive"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Find:"
         Height          =   195
         Index           =   7
         Left            =   1440
         TabIndex        =   32
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "[Start]:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.Frame fraInStr 
      Caption         =   " InStr Function "
      Height          =   1815
      Left            =   3480
      TabIndex        =   17
      Top             =   1680
      Width           =   3255
      Begin VB.CheckBox chkCaseInStr 
         Caption         =   "Case sensitive"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdOKInstr 
         Caption         =   "OK"
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtStartInStr 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtInStr 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Text            =   "AbcdefGhiJklmnoPqrsTuVWXyZ"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtFindInStr 
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Text            =   "hijklm"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Find:"
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   22
         Top             =   840
         Width           =   360
      End
   End
   Begin VB.Frame fraMid 
      Caption         =   " Mid Function "
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   3255
      Begin VB.TextBox txtMidLen 
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtMid 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Text            =   "1234567890"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtMidStart 
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdOKMid 
         Caption         =   "OK"
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "[length]:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1800
         TabIndex        =   15
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   420
      End
   End
   Begin VB.Frame fraRight 
      Caption         =   " Right Function "
      Height          =   1335
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtRight 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Text            =   "1234567890"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtRightLen 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdOKRight 
         Caption         =   "OK"
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Length:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   555
      End
   End
   Begin VB.Frame fraLeft 
      Caption         =   " Left Function "
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdOKLeft 
         Caption         =   "OK"
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtLeftLen 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "1234567890"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Length:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOKAppend_Click()
    Dim bytTest() As Byte
    
    'Convert the test string to a byte array.
    'We assume we already started with a byte array to begin with.
    bytTest() = StrConv(txtConcat.Text, vbFromUnicode)
    
    'Append a string to bytTest().
    bsAppendToByte bytTest(), txtAppend.Text
    
    'Convert the data back to string and display it.
    txtConcat.Text = StrConv(bytTest(), vbUnicode)
End Sub

Private Sub cmdOKInstr_Click()
    Dim bytTest() As Byte, lonRet As Long
    Dim lonStart As Long
    
    'If the user didn't enter a valid start position (or one at all), then set one.
    If IsNumeric(txtStartInStr.Text) = False Then
        lonStart = -1
    Else
        lonStart = CLng(txtStartInStr.Text)
    End If
    
    'Convert the test string to a byte array.
    'We assume we already started with a byte array to begin with.
    bytTest() = StrConv(txtInStr.Text, vbFromUnicode)
    
    'Perform the InStr() function on the byte array, and return as a string.
    lonRet = bsInStrString(lonStart, bytTest(), txtFindInStr.Text, chkCaseInStr.Value)
    
    MsgBox lonRet
    
    'If the text was found, highlight it in the TextBox.
    If lonRet > 0 Then
        
        With txtInStr
            .SetFocus
            .SelStart = lonRet - 1
            .SelLength = Len(txtFindInStr.Text)
        End With
    
    End If
    
End Sub

Private Sub cmdOKLeft_Click()
    Dim bytTest() As Byte, strRet As String
    
    'Convert the test string to a byte array.
    'We assume we already started with a byte array to begin with.
    bytTest() = StrConv(txtLeft.Text, vbFromUnicode)
    
    'Perform the Left() function on the byte array, and return as a string.
    strRet = bsStrLeft(bytTest(), CLng(txtLeftLen.Text))
    
    'Title of message box is just a check to make sure length returned is correct.
    'Both numbers should be the same.
    MsgBox strRet, vbInformation, Len(strRet) & "/" & txtLeftLen.Text
End Sub

Private Sub cmdOKMid_Click()
    Dim bytTest() As Byte, strRet As String
    Dim lonLength As Long
    
    'If the user didn't enter a valid length (or one at all), then set one.
    If IsNumeric(txtMidLen.Text) = False Then
        lonLength = -1
    Else
        lonLength = CLng(txtMidLen.Text)
    End If
    
    'Convert the test string to a byte array.
    'We assume we already started with a byte array to begin with.
    bytTest() = StrConv(txtMid.Text, vbFromUnicode)
    
    'Perform the Mid() function on the byte array, and return as a string.
    strRet = bsStrMid(bytTest(), CLng(txtMidStart.Text), lonLength)
    
    MsgBox strRet, vbInformation, Len(strRet)
End Sub

Private Sub cmdOKNum_Click()
    Dim bytTest() As Byte
    
    'Convert the test string to a byte array.
    'We assume we already started with a byte array to begin with.
    bytTest() = StrConv(txtIsNum.Text, vbFromUnicode)
    
    'Perform the IsNumeric() function on the byte array.
    MsgBox bsIsNumeric(bytTest())
End Sub

Private Sub cmdOKRev_Click()
    Dim lonStart As Long, lonRet As Long
    Dim bytTest() As Byte
    
    'Convert the test string to a byte array.
    'We assume we already started with a byte array to begin with.
    bytTest() = StrConv(txtInStrRev.Text, vbFromUnicode)
    
    'If they didn't enter a start value, or it's wrong, then just set it to 0.
    'The function will automatically make it start at the last character.
    If IsNumeric(txtStartRev.Text) = False Then
        lonStart = 0
    Else
        lonStart = CLng(txtStartRev.Text)
    End If
    
    'Perform the InStrRev() function on the byte array.
    lonRet = bsInStrRev(lonStart, bytTest(), txtFindRev.Text, chkCaseRev.Value)
    
    MsgBox lonRet
    
    If lonRet > 0 Then
        'Found the string. Highlight the string that was found.
        With txtInStrRev
            .SetFocus
            .SelStart = (lonRet - 1)
            .SelLength = Len(txtFindRev.Text)
        End With
    
    End If
    
End Sub

Private Sub cmdOKRight_Click()
    Dim bytTest() As Byte, strRet As String
    
    'Convert the test string to a byte array.
    'We assume we already started with a byte array to begin with.
    bytTest() = StrConv(txtRight.Text, vbFromUnicode)
    
    'Perform the Right() function on the byte array, and return as a string.
    strRet = bsStrRight(bytTest(), CLng(txtRightLen.Text))
    
    'Title of message box is just a check to make sure length returned is correct.
    'Both numbers should be the same.
    MsgBox strRet, vbInformation, Len(strRet) & "/" & txtRightLen.Text
End Sub
