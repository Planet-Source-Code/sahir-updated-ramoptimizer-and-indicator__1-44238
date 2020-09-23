VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "RAM Optimizer"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5400
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Optimize"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProcessIndicator 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   2535
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4471
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   2535
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4471
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "sksahir@yahoo.com"
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   3240
      Width           =   1440
   End
   Begin VB.Label lblavailableRAM 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   2460
      Width           =   480
   End
   Begin VB.Label LblMaxPhyicalRam 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   2460
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declaration for memory allocation and release
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private lastpcent As Single, lastTot As Long

'declaration for memory status
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

'type declaration required to use GlobalMemoryStatus api

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
'declaration to use in form to know the RAM Status
Private memoryInfo As MEMORYSTATUS

Private Sub Command1_Click()
'if there is error in memory allocation or release
On Error GoTo finish

Dim I As Integer
Dim Ptr_M() As Long
Dim PTR_F() As Long

'declaration used for  process indicator
Dim NF As Integer
Dim NM As Integer
Dim Rtn As Long
Dim Intact As Boolean
Dim MAX As Long
ProcessIndicator.Visible = True
ProcessIndicator.MAX = 32000
ProcessIndicator.Value = 0


    Do While Intact = False
         ProcessIndicator.Value = ProcessIndicator.Value + 1
         ReDim Preserve Ptr_M(NM)
         ReDim Preserve PTR_F(NF)
         
         
            Ptr_M(NM) = GlobalAlloc(&H2, 40000)
            PTR_F(NF) = GlobalAlloc(&H0, 40000)
         
         If Ptr_M(NM) = 0 Then Intact = True
         If PTR_F(NF) = 0 Then Intact = True
         NM = NM + 1
         NF = NF + 1
    Loop


finish:
'unload any allocated memory

ProcessIndicator.MAX = NM
ProcessIndicator.Value = 0
For I = NM - 1 To 0 Step -1
     
    If ProcessIndicator.Value = ProcessIndicator.MAX - 1 Then Else ProcessIndicator.Value = ProcessIndicator.Value + 1
    Rtn = GlobalFree(Ptr_M(I))
    ReDim Preserve Ptr_M(I)
    Rtn = GlobalFree(PTR_F(I))
    DoEvents
Next I

ProcessIndicator.Visible = False
End Sub
Function GetMemoryStatus()
  'get memory info
  DoEvents
  GlobalMemoryStatus memoryInfo
    
  Totp1 = Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10
  Availp1 = Int(memoryInfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10
  pcent = Int(Availp1 / Totp1 * 100)
  
'code for available RAM progressbar
ProgressBar1.Value = pcent
lblavailableRAM.Caption = Format(Availp1) & " MB"

'code for total RAM progressbar
Dim totalRam As Long
totalRam = Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10
LblMaxPhyicalRam = totalRam & " MB"
ProgressBar2.Value = 100
End Function



Private Sub Timer1_Timer()
GetMemoryStatus
End Sub
