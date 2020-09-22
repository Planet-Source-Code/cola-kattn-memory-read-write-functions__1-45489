Attribute VB_Name = "Module1"
'|--------------------------------------------------------------------
'| These functions are written by Cola-kattn (opersvik@hotmail.com)
'|--------------------------------------------------------------------
'| You can use the functions in this module for learning purpose or you can
'| use this whole module or part of it in your own trainers if you like, there
'| are no real copyright on the code :)
'| Enjoy!
'|
'| The usage of the functions is pretty straight forward.
'| If the Read functions are succesfull, they will return the found value in a specified memory
'| offset from a specified process
'|
'| The Write functions will return True if writing to the specified memory offset was succesfull
'| and False if it failed.
'| You can take use of these returns as error checking though I made error checking in every
'| function that will display an message box and terminate the function when an error occurs
'| (If the process window is not found, or the program can't get a process handle.)
'|
'| The memory offset is commonly found by using a memory search tool, such as GameHack or TSearch.
'| The windowname of the game\program can be found in the window list box when you hit Ctrl + Alt + Del
'|---------------------------------------------------------------------


'//All API declarations we will need to make these functions useful:

'Thanks to Robert Meffe for pointing out this API line because he didn't get it
'to work properly in his Win XP. Greets!
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Private Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

'||-------------------------------------------------------------------------------------------------||
'|| The two next functions read\write BYTE values.                                                  ||
'|| BYTE is an 8-bit datatype that can store values from 0 to 255.                                  ||
'||-------------------------------------------------------------------------------------------------||

Public Function ReadByte(Offset As Long, WindowName As String) As Byte

    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    Dim Value As Byte
    
    'Try to find the window that was passed in the variable WindowName to this function.
    hwnd = FindWindow(vbNullString, WindowName)
    
    If hwnd = 0 Then
        
        'This is executed if the window cannot be found.
        'You can add or write own code here to customize your program.
        
        MsgBox "Could not find process window!", vbCritical, "Read error"
        
        Exit Function
    
    End If
    
    'Get the window's process ID.
    GetWindowThreadProcessId hwnd, ProcessID
    
    'Get a process handle.
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then
        
        'This is executed if a process handle cannot be found.
        'You can add or write your own code here to customize your program.
        
        MsgBox "Could not get a process handle!", vbCritical, "Read error"
        
        Exit Function
    
    End If

    
    'Read a BYTE value from the specified memory offset.
    ReadProcessMem ProcessHandle, Offset, Value, 1, 0&
    
    'Return the found memory value.
    ReadByte = Value
    
    'It is important to close the current process handle.
    CloseHandle ProcessHandle
           
End Function

Public Function WriteByte(Offset As Long, WindowName As String, Value As Byte) As Boolean

    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    
    'Try to find the window that was passed in the variable WindowName to this function.
    hwnd = FindWindow(vbNullString, WindowName)
    
    If hwnd = 0 Then
        
        'This is executed if the window cannot be found.
        'You can add or write own code here to customize your program.
        
        MsgBox "Could not find process window!", vbCritical, "Write error"
        
        Exit Function
    
    End If
    
    'Get the window's process ID.
    GetWindowThreadProcessId hwnd, ProcessID
    
    'Get a process handle.
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then
        
        'This is executed if a process handle cannot be found.
        'You can add or write your own code here to customize your program.
        
        MsgBox "Could not get a process handle!", vbCritical, "Write error"
        
        Exit Function
    
    End If
    
    'Write a specified BYTE value to the specified memory offset.
    WriteProcessMemory ProcessHandle, Offset, Value, 1, 0&
    
    'It is important to close the current process handle.
    CloseHandle ProcessHandle
    
End Function


'||-------------------------------------------------------------------------------------------------||
'|| The two next functions read\write INTEGER values.                                               ||
'|| INTEGER is a 16-bit(2 byte) datatype and can store values from -32768 to 32767                  ||
'||-------------------------------------------------------------------------------------------------||

Public Function ReadInteger(Offset As Long, WindowName As String) As Integer

    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    Dim Value As Integer
    
    'Try to find the window that was passed in the variable WindowName to this function.
    hwnd = FindWindow(vbNullString, WindowName)
    
    If hwnd = 0 Then
        
        'This is executed if the window cannot be found.
        'You can add or write your own code here to customize your program.
        
        MsgBox "Could not find process window!", vbCritical, "Read error"
            
        Exit Function
    
    End If
    
    'Get the window's process ID.
    GetWindowThreadProcessId hwnd, ProcessID
    
    'Get a process handle.
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then

        'This is executed if a process handle cannot be found.
        'You can add or write your own code here to customize your program.

        MsgBox "Could not get a process handle!", vbCritical, "Read error"
        
        Exit Function
        
    End If
    
    'Read an INTEGER value from the specified memory offset.
    ReadProcessMem ProcessHandle, Offset, Value, 2, 0&
    
    'Return the found memory value.
    ReadInteger = Value
    
    'It is important to close the current process handle.
    CloseHandle ProcessHandle
    
End Function

Public Function WriteInteger(Offset As Long, WindowName As String, Value As Integer) As Boolean

    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    
    'Try to find the window that was passed in the variable WindowName to this function.
    hwnd = FindWindow(vbNullString, WindowName)
    
    If hwnd = 0 Then
    
        'This is executed if the window cannot be found.
        'You can add or write your own code here to customize your program.
        
        MsgBox "Could not find process window!", vbCritical, "Write error"
        
        Exit Function
    
    End If
    
    'Get the window's process ID.
    GetWindowThreadProcessId hwnd, ProcessID
    
    'Get a process handle.
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then
    
        'This is executed if a process handle cannot be found.
        'You can add or write your own code here to customize your program.
    
        MsgBox "Could not get a process handle!", vbCritical, "Write error"
        
        Exit Function
        
    End If
    
    'Write a specified INTEGER value to the specified memory offset.
    WriteProcessMemory ProcessHandle, Offset, Value, 2, 0&
    
    'It is important to close the current process handle.
    CloseHandle ProcessHandle
    
End Function


'||-------------------------------------------------------------------------------------------------||
'|| The two next functions read\write LONG values.                                                  ||
'|| LONG is a 32-bit(4 byte) datatype and can store values from -2,147,483,648 to 2,147,483,647     ||
'||-------------------------------------------------------------------------------------------------||

Public Function ReadLong(Offset As Long, WindowName As String) As Long

    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    Dim Value As Long
    
    'Try to find the window that was passed in the variable WindowName to this function.
    hwnd = FindWindow(vbNullString, WindowName)
    
    If hwnd = 0 Then
    
            'This is executed if the window cannot be found.
            'You can add or write your own code here to customize your program.
                        
            MsgBox "Could not find process window!", vbCritical, "Read error"
            
            Exit Function
        
    End If
    
    'Get the window's process ID.
    GetWindowThreadProcessId hwnd, ProcessID
    
    'Get a process handle
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then
    
        'This is executed if a process handle cannot be found.
        'You can add or write your own code here to customize your program.
    
        MsgBox "Could not get a process handle!", vbCritical, "Read error"
        
        Exit Function
        
    End If
    
    'Read a LONG from the specified memory offset.
    ReadProcessMem ProcessHandle, Offset, Value, 4, 0&
    
    'Return the found memory value.
    ReadLong = Value
    
    'It is important to close the current process handle.
    CloseHandle ProcessHandle
    
End Function

Public Function WriteLong(Offset As Long, WindowName As String, Value As Long) As Boolean

    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    
    'Try to find the window that was passed in the variable WindowName to this function.
    hwnd = FindWindow(vbNullString, WindowName)
    
    If hwnd = 0 Then
    
            'This is executed if the window cannot be found.
            'You can add or write your own code here to customize your program.
                        
            MsgBox "Could not find process window!", vbCritical, "Write error"
            
            Exit Function
        
    End If
    
    'Get the window's process ID.
    GetWindowThreadProcessId hwnd, ProcessID
    
    'Get a process handle
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then
    
        'This is executed if a process handle cannot be found.
        'You can add or write your own code here to customize your program.
    
        MsgBox "Could not get a process handle!", vbCritical, "Write error"
        
        Exit Function
        
    End If
    
    'Read a LONG from the specified memory offset.
    WriteProcessMemory ProcessHandle, Offset, Value, 4, 0&

    'It is important to close the current process handle.
    CloseHandle ProcessHandle

End Function

