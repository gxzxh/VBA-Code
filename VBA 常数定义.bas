Attribute VB_Name = "常数定义"
Option Explicit
Public Enum BinaryBit
    vbbit0 = 1
    vbbit1 = 2
    vbbit2 = 4
    vbbit3 = 8
    vbBit4 = &H10
    vbBit5 = &H20
    vbBit6 = &H40
    vbBit7 = &H80
    vbBit8 = &H100
    vbBit9 = &H200
    vbBit10 = &H400
    vbBit11 = &H800
    vbBit12 = &H1000
    vbBit13 = &H2000
    vbBit14 = &H4000
    vbBit15 = &H8000
    vbBit16 = &H10000
    vbBit17 = &H20000
    vbBit18 = &H40000
    vbBit19 = &H80000
    vbBit20 = &H100000
    vbBit21 = &H200000
    vbBit22 = &H400000
    vbBit23 = &H800000
    vbBit24 = &H1000000
    vbBit25 = &H2000000
    vbBit26 = &H4000000
    vbBit27 = &H8000000    '134217728
    vbBit28 = &H10000000    ' 268435456
    vbBit29 = &H20000000    '536870912
    vbBit30 = &H40000000    '1073741824
    vbBit31 = &H80000000
End Enum
Public Const DRIVE_REMOVABLE = 2    '表示软盘
Public Const DRIVE_FIXED = 3    '表示硬盘驱动器
Public Const DRIVE_REMOTE = 4    '表示网络驱动器
Public Const DRIVE_CDROM = 5    '表示光盘驱动器
Public Const DRIVE_RAMDISK = 6    '表示RAM驱动器
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Sub main()
    MsgBox vbBit27
End Sub

