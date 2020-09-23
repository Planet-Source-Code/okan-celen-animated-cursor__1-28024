VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   2580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Restore"
      Height          =   555
      Left            =   1140
      TabIndex        =   1
      Top             =   570
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code to be inserted into a module

Private Declare Function LoadCursorFromFile Lib "user32" Alias _
    "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetSystemCursor Lib "user32" _
    (ByVal hcur As Long, ByVal id As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function CopyIcon Lib "user32" (ByVal hcur As Long) As Long

Private Const OCR_NORMAL = 32512

Public lngOldCursor As Long, lngNewCursor As Long

Public Sub StartAnimatedCursor(AniFilePath As String)

    'Create a copy of the current cursor,
    'for Windows NT compatibility
    
    lngOldCursor = CopyIcon(GetCursor())
    
    'Check the passed string, if it contains
    'a solid file path, then load the cursor
    'from file. If not, add the App.Path,
    '*then* load cursor...
    
    If InStr(1, AniFilePath, "\") Then
        lngNewCursor = LoadCursorFromFile(AniFilePath)
    Else
        lngNewCursor = LoadCursorFromFile(App.Path & _
            "\" & AniFilePath)
    End If
    
    'Activate the cursor
        
    SetSystemCursor lngNewCursor, OCR_NORMAL
    
End Sub

Public Sub RestoreLastCursor()

    'Restore last cursor
    
    SetSystemCursor lngOldCursor, OCR_NORMAL

End Sub


Private Sub Command1_Click()
StartAnimatedCursor ("c:\windows\cursors\globe.ani")
End Sub

Private Sub Command2_Click()
RestoreLastCursor
End Sub

