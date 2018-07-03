VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.ListBox List1 
         BackColor       =   &H00FFC0C0&
         Height          =   1230
         ItemData        =   "frmMain.frx":08CA
         Left            =   120
         List            =   "frmMain.frx":08CC
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtBoardName 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdCheckBoardName 
         BackColor       =   &H00C0C0FF&
         Caption         =   "LoadBoard"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4320
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtBoardXY 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "c:\board_xy"
         ToolTipText     =   "Format file"
         Top             =   240
         Width           =   5655
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   1560
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strMBboardName As String
Dim bPanelBoard As Boolean
Private Sub cmdCheckBoardName_Click()
strMBboardName = ""
List1.Clear
bPanelBoard = False
cmdCheckBoardName.Enabled = False
Call ReadBoard_xy_ver
txtBoardName.Text = List1.List(0)
strMBboardName = Trim(txtBoardName.Text)
If strMBboardName = "" Then
  cmdGo.Enabled = False
  cmdCheckBoardName.Enabled = True
  Else
   cmdGo.Enabled = True
End If
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub cmdGo_Click()
Frame1.Enabled = False
 Call ReadBoard_xy
 Frame1.Enabled = True
End Sub

Private Sub cmdOpen_Click()
On Error GoTo errh
  Me.cdg.DialogTitle = "Choose board_xy file"
  Me.cdg.Filter = "board_xy|board_xy|*.*|*.*"
  Me.cdg.FileName = "board_xy"
  Me.cdg.ShowOpen
  If Me.cdg.FileName <> "" And LCase(Me.cdg.FileName) <> "board_xy" Then
     Me.txtBoardXY.Text = Trim(Me.cdg.FileName)
     If Dir(Me.txtBoardXY.Text) = "" Then
         Me.txtBoardXY.Text = "board_xy file not find"
         Exit Sub
     End If
      
     Else
        If Dir(Me.txtBoardXY.Text) = "" Then
            Me.txtBoardXY.Text = "board_xy file not find"
             Exit Sub
        End If
     
  End If
  cmdCheckBoardName.Enabled = True '
  
errh:
End Sub
Private Sub ReadBoard_xy()
Dim Mystr As String
Dim intI As Integer
Dim MystrTmp As String

Dim strTmpMB() As String
Dim bNode As Boolean
Dim strNodeName As String
Dim bBoardVer As Boolean
Dim bFixedOK As Boolean
Dim bStartFind As Boolean
    On Error Resume Next

 Open App.Path & "\Fixed_Node.txt" For Output As #4
 Open App.Path & "\Family_Fixed_Node.txt" For Output As #5
    Open txtBoardXY.Text For Input As #3
        Print #5, "FIXED NODE OPTIONS"
        
        Do Until EOF(3)
           Line Input #3, Mystr
           Dim strDevice() As String
           MystrTmp = UCase(Trim(Mystr))
        If MystrTmp <> "" Then
               'BOARDS
               If Left(MystrTmp, 6) = "BOARDS" Then bBoardVer = True
               
               If Left(MystrTmp, 6) <> "BOARDS" And bBoardVer = True And Left(MystrTmp, 2) = "1 " Then
                  strTmpMB = Split(MystrTmp, " ")
                  If strMBboardName = "" Then
                     strMBboardName = strTmpMB(1)
                     MsgBox "The board name is default MB(1%XXX)! " & strMBboardName
                  End If
                  bBoardVer = False
               End If
               'board
               If Left(MystrTmp, 5) = "OTHER" Then
                  ' Exit Do
                     If bStartFind = True Then
                       bStartFind = False
                       Exit Do
                     End If
               End If
               'OTHER
               
               If Left(MystrTmp, 6) = "BOARD " And InStr(MystrTmp, strMBboardName) <> 0 Then
                    
                    bStartFind = True
               End If
  '#################################
          If strMBboardName = "NonPanleBoard" And bPanelBoard = True Then
               bStartFind = True
          End If
  '################################
            'If Left(MystrTmp, 9) <> "!!!!   15" Then GoTo errh
            
            'node
            If Left(MystrTmp, 4) = "END " Then bStartFind = False
 If bStartFind = True And Left(MystrTmp, 6) <> "BOARD " Then
 'start true
            If Left(MystrTmp, 5) = "NODE " Then
                strDevice = Split(MystrTmp, " ")
                bNode = True
                intBoardAllNet = intBoardAllNet + 1
                strNodeName = Trim(UCase(strDevice(1)))
                If strNodeName <> "" Then
                   '+
                   If Left(strNodeName, 1) = "+" Then
                      
                      bFixedOK = True
                   End If
                   
                   If InStr(strNodeName, "5V") <> 0 Or InStr(strNodeName, "3V") <> 0 Or InStr(strNodeName, "3D3V") <> 0 Then
                      
                      bFixedOK = True
                   End If
                    If InStr(strNodeName, "_PWR") <> 0 Or InStr(strNodeName, "_ALW") <> 0 Or InStr(strNodeName, "_SUS") <> 0 Then
                      
                       bFixedOK = True
                   End If
                    If InStr(strNodeName, "_S3") <> 0 Or InStr(strNodeName, "_S0") <> 0 Or InStr(strNodeName, "_S0") <> 0 Then
                      
                       bFixedOK = True
                   End If
                   If InStr(strNodeName, "_VCC_") <> 0 Or InStr(strNodeName, "VCC_") <> 0 Or InStr(strNodeName, "_VTT") <> 0 Then
                      
                       bFixedOK = True
                   End If
                   If InStr(strNodeName, "_SLP_") Or InStr(strNodeName, "_+") Then
                      
                       bFixedOK = True
                   End If
                   
                   
                   If strNodeName = "LCDVDD" Then
                        bFixedOK = True
                   End If
                    If strNodeName = "DCBATOUT" Then
                        bFixedOK = True
                   End If
                   If strNodeName = "AD+" Or strNodeName = "VBAT" Or strNodeName = "CPU_CORE" Or strNodeName = "BAT" Or strNodeName = "BT+" Then
                        bFixedOK = True
                   End If
                    If InStr(strNodeName, "VDD_PMU_LDO") <> 0 Then
                        bFixedOK = True
                   End If
                   
                   '
                    If InStr(strNodeName, "PWR_1D05V_") <> 0 Or InStr(strNodeName, "_FB") <> 0 Or InStr(strNodeName, "PWR_5V_") <> 0 Or InStr(strNodeName, "PWR_3D3V_") <> 0 Or InStr(strNodeName, "PWR_1D5V_") <> 0 Then
                      
                      bFixedOK = False
                   End If
                   
                   'PIN BAN
                   '
                    If InStr(strNodeName, "P_+") <> 0 Or InStr(strNodeName, "SCL") <> 0 Or InStr(strNodeName, "SDA") <> 0 Then
                        bFixedOK = False
                    End If
                End If
                If bFixedOK = True Then
                   Print #4, strNodeName
                   Print #5, Tab(4); strNodeName & " Family ALL is 1;"
                End If
                
                
                bFixedOK = False
                
                 strNodeName = ""
                    
             

             End If
         End If
 End If 'start true
            
          intI = intI + 1
          Me.Caption = intI
            DoEvents
          
            
        Loop
         Print #5, Tab(4); "gnd Family ALL is 0;"
         Print #5, Tab(4); "gnd GROUND;"
    Close #4
     Close #3
    Close #5
    Me.Caption = "File Output Ok!"
    MsgBox "File Output Ok!"
    Exit Sub
    
errh:
 MsgBox "The board_xy file error,please check!", vbCritical
End Sub
Private Sub ReadBoard_xy_ver()


Dim Mystr As String
Dim intI As Integer
Dim MystrTmp As String
Dim bBoardVer As Boolean
Dim strTmpMB() As String
bPanelBoard = False
 intI = 0
    On Error Resume Next
    Open txtBoardXY.Text For Input As #23
        Do Until EOF(23)
           Line Input #23, Mystr
           MystrTmp = UCase(Trim(Mystr))
        If MystrTmp <> "" Then
               'BOARDS
               If Left(MystrTmp, 5) = "BOARD" And Left(MystrTmp, 6) <> "BOARDS" Then Exit Do
               
               If Left(MystrTmp, 6) = "BOARDS" Then bBoardVer = True
               If Left(MystrTmp, 6) <> "BOARDS" And bBoardVer = True Then ' And Left(MystrTmp, 2) = "1 " Then
                  strTmpMB = Split(MystrTmp, " ")
                  strMBboardName = strTmpMB(1)
                  List1.List(intI) = strTmpMB(1)
                  'bBoardVer = False
                  intI = intI + 1
               End If
         End If
        Loop
        DoEvents
    Close #23
      If List1.List(0) = "" Then
      
                  strMsg = MsgBox("The board not is Panel Board!,Do you want to continue ?", 52, "Warning!")
            If strMsg = vbYes Then
                  GoTo START
               ElseIf strMsg = vbNo Then
               bPanelBoard = False
                Exit Sub
            End If

   
      End If
      
   Exit Sub
START:
     List1.List(0) = "NonPanleBoard"
      bPanelBoard = True
            
      
      
End Sub
