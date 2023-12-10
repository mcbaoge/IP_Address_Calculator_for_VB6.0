VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   7140
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   12
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox subnetmask 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox IPaddress4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox IPaddress3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox IPaddress2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox IPaddress1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label5 
      Height          =   1335
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label results 
      Caption         =   "计算结果："
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "IP地址："
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function DecimalToBinary(ByVal decimalNumber As Long) As String

    Dim binaryResult As String
    
    Do
    
        binaryResult = CStr(decimalNumber Mod 2) & binaryResult
        
        decimalNumber = decimalNumber \ 2
        
    Loop Until decimalNumber = 0
    
    DecimalToBinary = binaryResult
    
End Function

Function BinaryToDecimal(ByVal binaryString As String) As Long

    Dim decimalNumber As Long
    
    Dim i As Integer
    
    Dim length As Integer
    
    length = Len(binaryString)
    
    For i = 1 To length
    
        Dim digit As Integer
        
        digit = CInt(Mid(binaryString, length - i + 1, 1))
        
        decimalNumber = decimalNumber + (digit * (2 ^ (i - 1)))
        
    Next i
    
    BinaryToDecimal = decimalNumber
    
End Function

Private Sub Command1_Click()

    If subnetmask.Text >= 24 And subnetmask.Text < 32 Then
    
        network_number = subnetmask.Text - 24
        
        Level = 4
        
    ElseIf subnetmask.Text >= 16 And subnetmask.Text < 24 Then
    
        network_number = subnetmask.Text - 16
        
        Level = 3
        
    ElseIf subnetmask.Text >= 8 And subnetmask.Text < 16 Then
    
        network_number = subnetmask.Text - 8
        
        Level = 2
        
    Else
    
        network_number = subnetmask.Text
        
        Level = 1
        
    End If
    
    For i = 1 To 8 Step 1
    
        If Len(subnetmasknumber) <= network_number - 1 Then
    
            subnetmasknumber = subnetmasknumber & "1"
            
        Else
        
            subnetmasknumber = subnetmasknumber & "0"
            
        End If
        
    Next i
    
    If Level = 4 Then
    
        '第四部分的IP地址处理
        '算网络地址
        binaryString = DecimalToBinary(IPaddress4)
        
        address = "" & binaryString
        
        Do
        
            If Len(address) <> 8 Then
            
                address = "0" & address
                
            Else
            
                Exit Do
                
            End If
        
        Loop
    
        For i = 1 To 8
        
            jisuan = Val(Mid(address, i, 1)) And Val(Mid(subnetmasknumber, i, 1))
            
            erjinzhijieguo = erjinzhijieguo & jisuan
            
        Next i
        
        decimalNumber = BinaryToDecimal(Val(erjinzhijieguo))
        '算广播地址
        count1 = 0
        
        For i = 1 To Len(subnetmasknumber)
        
            If Mid(subnetmasknumber, i, 1) = "1" Then
            
                count1 = count1 + 1
                
            End If
            
        Next i
        
        a = Left(address, count1)
        
        Do
        
            If Len(a) < 8 Then
            
                a = a & "1"
                
            Else
            
                Exit Do
                
            End If
            
        Loop
        
        decimalNumber1 = BinaryToDecimal(Val(a))
        
        jieguo1 = "网络号：" & IPaddress1 & "." & IPaddress2 & "." & IPaddress3 & "." & decimalNumber
        
        jieguo2 = "第一个可用地址：" & IPaddress1 & "." & IPaddress2 & "." & IPaddress3 & "." & decimalNumber + 1
        
        jieguo3 = "最后一个可用地址：" & IPaddress1 & "." & IPaddress2 & "." & IPaddress3 & "." & decimalNumber1 - 1
        
        jieguo4 = "广播地址：" & IPaddress1 & "." & IPaddress2 & "." & IPaddress3 & "." & decimalNumber1
        
        jieguo = jieguo1 & vbCrLf & jieguo2 & vbCrLf & jieguo3 & vbCrLf & jieguo4
        
    ElseIf Level = 3 Then
        '第三部分的IP地址处理
        '算网络地址
        binaryString = DecimalToBinary(IPaddress3)
        
        address = "" & binaryString
        
        Do
        
            If Len(address) <> 8 Then
            
                address = "0" & address
                
            Else
            
                Exit Do
                
            End If
        
        Loop
    
        For i = 1 To 8
        
            jisuan = Val(Mid(address, i, 1)) And Val(Mid(subnetmasknumber, i, 1))
            
            erjinzhijieguo = erjinzhijieguo & jisuan
            
        Next i
        
        decimalNumber = BinaryToDecimal(Val(erjinzhijieguo))
        '算广播地址
        count1 = 0
        
        For i = 1 To Len(subnetmasknumber)
        
            If Mid(subnetmasknumber, i, 1) = "1" Then
            
                count1 = count1 + 1
                
            End If
            
        Next i
        
        a = Left(address, count1)
        
        Do
        
            If Len(a) < 8 Then
            
                a = a & "1"
                
            Else
            
                Exit Do
                
            End If
            
        Loop
        
        decimalNumber1 = BinaryToDecimal(Val(a))
        
        jieguo1 = "网络号：" & IPaddress1 & "." & IPaddress2 & "." & decimalNumber & "." & "0"
        
        jieguo2 = "第一个可用地址：" & IPaddress1 & "." & IPaddress2 & "." & decimalNumber & "." & "1"
        
        jieguo3 = "最后一个可用地址：" & IPaddress1 & "." & IPaddress2 & "." & decimalNumber1 & "." & "254"
        
        jieguo4 = "广播地址：" & IPaddress1 & "." & IPaddress2 & "." & decimalNumber1 & "." & "255"
        
        jieguo = jieguo1 & vbCrLf & jieguo2 & vbCrLf & jieguo3 & vbCrLf & jieguo4
        
    ElseIf Level = 2 Then
        '第二部分的IP地址处理
        '算网络地址
        binaryString = DecimalToBinary(IPaddress2)
        
        address = "" & binaryString
        
        Do
        
            If Len(address) <> 8 Then
            
                address = "0" & address
                
            Else
            
                Exit Do
                
            End If
        
        Loop
    
        For i = 1 To 8
        
            jisuan = Val(Mid(address, i, 1)) And Val(Mid(subnetmasknumber, i, 1))
            
            erjinzhijieguo = erjinzhijieguo & jisuan
            
        Next i
        
        decimalNumber = BinaryToDecimal(Val(erjinzhijieguo))
        '算广播地址
        count1 = 0
        
        For i = 1 To Len(subnetmasknumber)
        
            If Mid(subnetmasknumber, i, 1) = "1" Then
            
                count1 = count1 + 1
                
            End If
            
        Next i
        
        a = Left(address, count1)
        
        Do
        
            If Len(a) < 8 Then
            
                a = a & "1"
                
            Else
            
                Exit Do
                
            End If
            
        Loop
        
        decimalNumber1 = BinaryToDecimal(Val(a))
        
        jieguo1 = "网络号：" & IPaddress1 & "." & decimalNumber & "." & "0" & "." & "0"
        
        jieguo2 = "第一个可用地址：" & IPaddress1 & "." & decimalNumber & "." & "0" & "." & "1"
        
        jieguo3 = "最后一个可用地址：" & IPaddress1 & "." & decimalNumber1 & "." & "255" & "." & "254"
        
        jieguo4 = "广播地址：" & IPaddress1 & "." & decimalNumber1 & "." & "255" & "." & "255"
        
        jieguo = jieguo1 & vbCrLf & jieguo2 & vbCrLf & jieguo3 & vbCrLf & jieguo4
        
        ElseIf Level = 1 Then
        '第一部分的IP地址处理
        '算网络地址
        binaryString = DecimalToBinary(IPaddress1)
        
        address = "" & binaryString
        
        Do
        
            If Len(address) <> 8 Then
            
                address = "0" & address
                
            Else
            
                Exit Do
                
            End If
        
        Loop
    
        For i = 1 To 8
        
            jisuan = Val(Mid(address, i, 1)) And Val(Mid(subnetmasknumber, i, 1))
            
            erjinzhijieguo = erjinzhijieguo & jisuan
            
        Next i
        
        decimalNumber = BinaryToDecimal(Val(erjinzhijieguo))
        '算广播地址
        count1 = 0
        
        For i = 1 To Len(subnetmasknumber)
        
            If Mid(subnetmasknumber, i, 1) = "1" Then
            
                count1 = count1 + 1
                
            End If
            
        Next i
        
        a = Left(address, count1)
        
        Do
        
            If Len(a) < 8 Then
            
                a = a & "1"
                
            Else
            
                Exit Do
                
            End If
            
        Loop
        
        decimalNumber1 = BinaryToDecimal(Val(a))
        
        jieguo1 = "网络号：" & decimalNumber & "." & "0" & "." & "0" & "." & "0"
        
        jieguo2 = "第一个可用地址：" & decimalNumber & "." & "0" & "." & "0" & "." & "1"
        
        jieguo3 = "最后一个可用地址：" & decimalNumber1 & "." & "255" & "." & "255" & "." & "254"
        
        jieguo4 = "广播地址：" & decimalNumber1 & "." & "255" & "." & "255" & "." & "255"
        
        jieguo = jieguo1 & vbCrLf & jieguo2 & vbCrLf & jieguo3 & vbCrLf & jieguo4
    
    End If
    
    Label5.Caption = jieguo

End Sub

Private Sub IPaddress1_Change()

    If Val(IPaddress1) > 255 Then
    
        MsgBox "你不能输入大于255的数值", vbExclamation, "警告"
        
        IPaddress1.Text = ""
    
    ElseIf Len(IPaddress1) = 3 Then
    
        IPaddress2.SetFocus
        
    End If

End Sub

Private Sub IPaddress2_Change()
    
    If Val(IPaddress2) > 255 Then
    
        MsgBox "你不能输入大于255的数值", vbExclamation, "警告"
        
        IPaddress2.Text = ""
    
    ElseIf Len(IPaddress2) = 3 Then
    
        IPaddress3.SetFocus
        
    End If

End Sub

Private Sub IPaddress3_Change()
    
    If Val(IPaddress3) > 255 Then
    
        MsgBox "你不能输入大于255的数值", vbExclamation, "警告"
        
        IPaddress3.Text = ""
    
    ElseIf Len(IPaddress3) = 3 Then
    
        IPaddress4.SetFocus
        
    End If

End Sub

Private Sub IPaddress4_Change()
    
    If Val(IPaddress4) > 255 Then
    
        MsgBox "你不能输入大于255的数值", vbExclamation, "警告"
        
        IPaddress4.Text = ""
    
    ElseIf Len(IPaddress4) = 3 Then
    
        subnetmask.SetFocus
        
    End If
    
End Sub

Private Sub subnetmask_Change()

    If Len(subnetmask.Text) = 2 Then
    
        If subnetmask.Text > 32 Then
        
            MsgBox "你不能输入大于32的子网掩码", vbExclamation, "警告"
            
            subnetmask.Text = ""
        
        End If
        
    End If

End Sub

Private Sub IPaddress1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' 允许退格键
            ' 继续允许输入
        Case 46 ' 小数点
            IPaddress2.SetFocus
            KeyAscii = 0 ' 阻止小数点显示在输入框中
        Case 48 To 57 ' 允许数字
            ' 继续允许输入
        Case Else
            KeyAscii = 0 ' 其他字符都不允许输入
    End Select
End Sub
Private Sub IPaddress2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' 允许退格键
            If Len(IPaddress2.Text) = 0 Then
                IPaddress3.SetFocus
            End If
        Case 46 '小数点
            IPaddress3.SetFocus
            KeyAscii = 0 ' 阻止小数点显示在输入框中
        Case 48 To 57 ' 允许数字
            ' 继续允许输入
        Case Else
            KeyAscii = 0 ' 其他字符都不允许输入
    End Select
End Sub
Private Sub IPaddress3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' 允许退格键
            If Len(IPaddress3.Text) = 0 Then
                IPaddress4.SetFocus
            End If
        Case 46 '小数点
            IPaddress4.SetFocus
            KeyAscii = 0 ' 阻止小数点显示在输入框中
        Case 48 To 57 ' 允许数字
            ' 继续允许输入
        Case Else
            KeyAscii = 0 ' 其他字符都不允许输入
    End Select
End Sub
Private Sub IPaddress4_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' 允许退格键
            If Len(IPaddress4.Text) = 0 Then
                subnetmask.SetFocus
            End If
        Case 46 '小数点
            subnetmask.SetFocus
            KeyAscii = 0 ' 阻止小数点显示在输入框中
        Case 48 To 57 ' 允许数字
            ' 继续允许输入
        Case Else
            KeyAscii = 0 ' 其他字符都不允许输入
    End Select
End Sub
Private Sub subnetmask_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 46 ' 允许退格键和小数点
            If Len(subnetmask.Text) = 0 Then
                IPaddress4.SetFocus
            End If
        Case 48 To 57 ' 允许数字
            ' 继续允许输入
        Case Else
            KeyAscii = 0 ' 其他字符都不允许输入
    End Select
End Sub
