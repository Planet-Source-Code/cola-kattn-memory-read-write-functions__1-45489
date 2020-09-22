VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test memory manipulation functions"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWlValue 
      Height          =   285
      Left            =   5160
      TabIndex        =   26
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtWlWindowName 
      Height          =   285
      Left            =   3480
      TabIndex        =   25
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtWlOffset 
      Height          =   285
      Left            =   1800
      TabIndex        =   24
      Text            =   "&H"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdWriteLong 
      Caption         =   "Write Long"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtRlValue 
      Height          =   285
      Left            =   5160
      TabIndex        =   22
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtRlWindowName 
      Height          =   285
      Left            =   3480
      TabIndex        =   21
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtRlOffset 
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Text            =   "&H"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdReadLong 
      Caption         =   "Read Long"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtWiValue 
      Height          =   285
      Left            =   5160
      TabIndex        =   18
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtWiWindowName 
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtWiOffset 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Text            =   "&H"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdWriteInteger 
      Caption         =   "Write Integer"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtRiValue 
      Height          =   285
      Left            =   5160
      TabIndex        =   14
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtRiWindowName 
      Height          =   285
      Left            =   3480
      TabIndex        =   13
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtRiOffset 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Text            =   "&H"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdReadInteger 
      Caption         =   "Read Integer"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtWbValue 
      Height          =   285
      Left            =   5160
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtWbWindowName 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtWbOffset 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   "&H"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdWriteByte 
      Caption         =   "Write Byte"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtRbValue 
      Height          =   285
      Left            =   5160
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtRbWindowName 
      Height          =   285
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtRbOffset 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "&H"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdReadByte 
      Caption         =   "Read Byte"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Value:"
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "WindowName:"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Offset:"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'|-------------------------------------------------------------------
'| This is a simple program to test the memory read\write functions by letting the
'| user read or write to a specified process' memory.
'|
'| When reading from a specified memory offset, the returned value will appear in
'| the value field, and when writing to a specified memory offset, you have to type
'| the value of your choice in the value field.
'| Remember that hex numbers in Visual Basic are identified by putting &H in front of the number.
'| This means that every offset must have &H in front of it.
'| Note that if you get a type mismatch error you have probably entered a too high/low number in the
'| value field (entering 500 in the byte or 35000 in the integer field for example) or
'| the memory offset might be in a wrong format (remember &H in front of the actual hex value)
'|
'| Written by Cola-Kattn. Hope you like it :)
'|---------------------------------------------------------------------


Private Sub cmdReadByte_Click()
    
    'Dimension the variable to store the value that the read function returns.
    Dim Value As Byte

    'Call the function to read a byte value from the specified memory offset and
    'store the returned value to the Value variable.
    Value = ReadByte(txtRbOffset.Text, txtRbWindowName.Text)
    
    'Add the returned value to the Read Byte value field.
    txtRbValue.Text = Value
    
End Sub


Private Sub cmdWriteByte_Click()
    
    'The variable must be a byte datatype here to match the WriteByte function arguments.
    Dim Value As Byte
    
    Value = txtWbValue.Text
    
    'Call the function that will write the specified value to the specified memory offset.
    WriteByte txtWbOffset.Text, txtWbWindowName.Text, Value
       
End Sub

Private Sub cmdReadInteger_Click()

    'Dimension a variable to store the value that the read function returns.
    Dim Value As Integer
    
    'Call the function to read an integer value from the specified memory offset and
    'store the returned value in Value variable.
    Value = ReadInteger(txtRiOffset.Text, txtRiWindowName)
    
    'Add the returned value to the Read Integer value field.
    txtRiValue.Text = Value

End Sub

Private Sub cmdWriteInteger_Click()

    'The value must be an integer here to match the WriteInteger function arguments.
    Dim Value As Integer
    
    Value = txtWiValue.Text
    
    'Call the function that will write the specified value to the specified memory offset.
    WriteInteger txtWiOffset.Text, txtWiWindowName.Text, Value

End Sub


Private Sub cmdReadLong_Click()
    
    'Dimension a variable to store the value the the read function returns.
    Dim Value As Long
    
    'Call the function to read a long value from the specified memory offset and
    'store the returned value in Value variable.
    Value = ReadLong(txtRlOffset.Text, txtRlWindowName.Text)
    
    'Add the returned value to the Read Long value field.
    txtRlValue.Text = Value

End Sub

Private Sub cmdWriteLong_Click()

    'Dimension the variable as long here to match the WriteLong function arguments.
    Dim Value As Long
    
    'Add the Write Long value field's value to the Value variable.
    Value = txtWlValue.Text
    
    'Call the function that will write the specified value to the specified memory offset.
    WriteLong txtWlOffset.Text, txtWlWindowName.Text, Value
    
End Sub
