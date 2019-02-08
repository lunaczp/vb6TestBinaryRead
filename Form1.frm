VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call UsingBytes()
   Call UsingString()
   DoEvents
End Sub

Private Sub UsingBytes()
   filename = "test.txt"

   Dim fileNum As Integer
   Dim bytes() As Byte

   fileNum = FreeFile
   Open Dir(filename) For Binary As fileNum
   ReDim bytes(1 To LOF(fileNum))
   Get fileNum, , bytes
   Close fileNum

   For i = LBound(bytes) To UBound(bytes)
      Debug.Print "Offset " & i & ": " & Hex(bytes(i))
   Next

   rst$ =""
   For i = LBound(bytes) To UBound(bytes)
      rst$ = rst$ & Chr(bytes(i))
   Next

   Debug.Print "byte result:" & rst
   Debug.Print "byte result len:" & Len(rst)

   For i = 1 To Len(rst)
      Debug.Print "reCheck Offset " & i & ": " & Hex(Asc(Mid$(rst, i, 1)))
   Next

End Sub

Private Sub UsingString()
   filename = "test.txt"

   Dim fileNum As Integer
   Dim fileContent As String

   fileNum = FreeFile
   Open Dir(filename) For Binary As fileNum
   fileContent = String$(LOF(fileNum), " ")
   Get fileNum, , fileContent
   Close fileNum

   For i = 1 To Len(fileContent)
      Debug.Print "Offset " & i & ": " & Hex(Asc(Mid$(fileContent, i, 1)))
   Next

   Debug.Print "string:result:" &fileContent
   Debug.Print "string:result len:" & Len(fileContent)

End Sub
