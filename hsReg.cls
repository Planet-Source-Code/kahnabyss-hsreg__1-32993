VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hsReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
(ByVal hKey As HKeys, ByVal lpSubKey As String, _
ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
(ByVal hKey As HKeys, ByVal lpSubKey As String, ByVal Reserved As Long, _
ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
lpdwDisposition As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
(ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
(ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As HKeys, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As HKeys, ByVal lpValueName As String, ByVal Reserved As Long, _
ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
(ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type GetRegValue
  ValueName As String * 255
  ValueNameL As Long
  ValueL As Long
  FirstChar As Byte
  OtherChar(251) As Byte
End Type

Private Const Kaa = &HF003F

'/---------------------------------------\
'| Description: Open a Registry Key      |
'| Input:       hKey System Key to open  |
'|              SubKey to open if needed |
'| Output:      Handle of the key or     |
'|              0 if function failed     |
'| Author:      KahnAbyss                |
'| Created:     22/03/02                 |
'| Modified:    22/03/02                 |
'\---------------------------------------/
Function OpenKey(hKey As HKeys, Optional SubKey As String) As Long
Dim RetVal As Long
Dim hndKey As Long
  
  RetVal = RegOpenKeyEx(hKey, SubKey, 0, Kaa, hndKey)
  If RetVal = 0 Then OpenKey = hndKey
End Function

'/---------------------------------------\
'| Description: Close an opened key      |
'| Input:       Handle of an opened key  |
'| Author:      KahnAbyss                |
'| Created:     22/03/02                 |
'| Modified:    22/03/02                 |
'\---------------------------------------/
Sub CloseKey(hndKey As Long)
  RegCloseKey hndKey
End Sub

'/---------------------------------------\
'| Description: Create a Registry Key    |
'| Input:       hndKey opened key        |
'|              SubKey to create         |
'| Output:      Handle of the key or     |
'|              0 if function failed     |
'| Author:      KahnAbyss                |
'| Created:     23/03/02                 |
'| Modified:    23/03/02                 |
'\---------------------------------------/
Function CreateKey(hndKey As Long, SubKey As String) As Long
Dim RetVal As Long
Dim hKey As Long
Dim KeyStatus As Long
Dim KeyClass As String * 255
Dim SAttr As SECURITY_ATTRIBUTES
  
  RetVal = RegCreateKeyEx(hndKey, SubKey, 0, KeyClass, _
                          0, Kaa, SAttr, hKey, KeyStatus)
  If RetVal = 0 Then CreateKey = hKey
End Function

'/---------------------------------------\
'| Description: Enumerate subkeys of an  |
'|              opened key               |
'| Input:       hndKey opened key        |
'| Output:      Subkeys name             |
'| Author:      KahnAbyss                |
'| Created:     23/03/02                 |
'| Modified:    23/03/02                 |
'\---------------------------------------/
Function EnumKey(hndKey As Long) As String()
Dim CurKey As Byte
Dim Done As Integer
Dim FTime As FILETIME
Dim RetVal() As GetRegValue
Dim Keys() As String
  
  Do
    ReDim Preserve RetVal(CurKey)
    With RetVal(CurKey)
      .ValueNameL = 255
      Done = RegEnumKeyEx(hndKey, CurKey, .ValueName, .ValueNameL, 0, _
            vbNull, vbNull, FTime)
      If Done = 259 Then Exit Do
      ReDim Preserve Keys(CurKey)
      Keys(CurKey) = Left(.ValueName, .ValueNameL)
    End With
    CurKey = CurKey + 1
  Loop
  EnumKey = Keys
End Function

'/---------------------------------------\
'| Description: Delete subkeys of an     |
'|              opened key               |
'| Input:       hndKey opened key        |
'|              subkey to delete         |
'| Output:      0 if function succeed    |
'| Author:      KahnAbyss                |
'| Created:     23/03/02                 |
'| Modified:    23/03/02                 |
'\---------------------------------------/
Function DeleteKey(hndKey As Long, SubKey As String) As Long

  DeleteKey = RegDeleteKey(hndKey, SubKey)
End Function

'/---------------------------------------\
'| Description: Return a value content   |
'|              of a registry value      |
'| Input:       hndKey opened key        |
'|              kValue return value of   |
'|              queried value            |
'|              If not specified return  |
'|              default value            |
'| Output:      Value Content            |
'| Author:      KahnAbyss                |
'| Created:     22/03/02                 |
'| Modified:    22/03/02                 |
'\---------------------------------------/
Function QueryValue(hndKey As Long, Optional kValue As String) As String
Dim RetVal As GetRegValue
  
  RetVal.ValueL = 255
  RegQueryValueEx hndKey, kValue, 0, vbNull, RetVal.ValueName, RetVal.ValueL
  QueryValue = Left(RetVal.ValueName, RetVal.ValueL - 1)
End Function

'/---------------------------------------\
'| Description: Save/create a value      |
'| Input:       hndKey opened key        |
'|              ValueContent content to  |
'|              insert into registry     |
'|              ValueName name of the    |
'|              value to set if nothing  |
'|              specified it will be the |
'|              default key              |
'| Output:      0 If function suceed     |
'| Author:      KahnAbyss                |
'| Created:     22/03/02                 |
'| Modified:    22/03/02                 |
'\---------------------------------------/
Function SetValue(hndKey As Long, ValueContent As String, Optional ValueName As String) As Long
  
  SetValue = RegSetValueEx(hndKey, ValueName, 0, 1, ValueContent, Len(ValueContent))
End Function

'/---------------------------------------\
'| Description: Enumerate Value of an    |
'|              opened key               |
'| Input:       hndKey opened key        |
'| Output:      Values() array of string |
'| Author:      KahnAbyss                |
'| Created:     22/03/02                 |
'| Modified:    22/03/02                 |
'\---------------------------------------/
Function EnumValue(hndKey As Long, ByRef Values() As hsRetVal)
Dim tmpValues() As New hsRetVal
Dim RetVal() As GetRegValue
Dim CurValue As Integer
Dim Done As Integer
  
  Do
    ReDim Preserve RetVal(CurValue)
    With RetVal(CurValue)
      .ValueNameL = 255
      .ValueL = 255
      Done = RegEnumValue(hndKey, CurValue, .ValueName, .ValueNameL, _
                          0, 0, .FirstChar, .ValueL)
      If Done = 259 Then Exit Do
      ReDim Preserve tmpValues(CurValue)
      tmpValues(CurValue).ValueName = Left(.ValueName, .ValueNameL)
      tmpValues(CurValue).ValueContent = Chr(.FirstChar)
      For X = 0 To .ValueL - 3
        tmpValues(CurValue).ValueContent = tmpValues(CurValue).ValueContent & _
        Chr(.OtherChar(X))
      Next X
    End With
    CurValue = CurValue + 1
  Loop
  Values = tmpValues
End Function

'/---------------------------------------\
'| Description: Delete value of an       |
'|              opened key               |
'| Input:       hndKey opened key        |
'|              value to delete          |
'| Output:      0 if function succeed    |
'| Author:      KahnAbyss                |
'| Created:     23/03/02                 |
'| Modified:    23/03/02                 |
'\---------------------------------------/
Function DeleteValue(hndKey As Long, kValue As String) As Long
  
  DeleteValue = RegDeleteValue(hndKey, kValue)
End Function
