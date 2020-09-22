VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Trainme"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtlong 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeLong 
      Caption         =   "Change"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Change long value."
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtInteger 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeInteger 
      Caption         =   "Change"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Change integer value."
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdChangeByte 
      Caption         =   "Change"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Change byte value."
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtByte 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer tmrExactValues 
      Interval        =   100
      Left            =   1560
      Top             =   840
   End
   Begin VB.Label Label4 
      Caption         =   "(Long)"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "(Integer)"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "(Byte)"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a small program which I wrote to test that all the memory read\write functions
'were working.
'
'Every variable that is initialized by the program gets a unique memory offset where its
'value is stored, the memory offsets for each variable in this program are in a text file named
'offsets.txt, but you can also use a memory search tool(such as GameHack or TSearch)
'to find the offsets if you like.

Dim ByteTest As Byte
Dim IntegerTest As Integer
Dim LongTest As Long

Private Sub Form_Load()
    
    'Set the Byte value to 100
    ByteTest = 100
    
    'Set the Integer value to 1600
    IntegerTest = 16000
    
    'Set the Long value to 120000000
    LongTest = 120000000
    
End Sub

Private Sub cmdChangeByte_Click()

    'When the first Change button is clicked, change the Byte value.
 
    Randomize Timer

    'This will generate a random Byte value from 0 to 255.
    ByteTest = Int(Rnd * 255)
    
End Sub

Private Sub cmdChangeInteger_Click()

    'If the first Change button is clicked, change the Integer value.

    Randomize Timer
    
    'This will generate a random integer value (Any number from -32768 to 32767)
    IntegerTest = Int((32767 - -32768 + 1) * Rnd + -32768)
    

End Sub

Private Sub cmdChangeLong_Click()

    'If the third Change button is clicked, change the Long value.

    Randomize Timer
    
    'This will generate a random Long value (Any number from -2147483648 to 2147483647)
    LongTest = Int((2147483647 - -2147483648#) * Rnd + -2147483648#)

End Sub


Private Sub tmrExactValues_Timer()
    
    'This timer will make sure every text contains the corresponding variable value.

    txtByte.Text = ByteTest

    txtInteger.Text = IntegerTest
    
    txtlong.Text = LongTest

End Sub

