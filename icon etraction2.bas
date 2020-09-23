Attribute VB_Name = "icons"
Option Explicit
'icon sizelocated in GetString(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", 32)
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

Public Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Public Enum iconsize
    largeIcon = 0
    SmallIcon = 1
End Enum

Public Const SH_USEFILEATTRIBUTES As Long = &H10
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const SHGFI_DISPLAYNAME  As Long = &H200
Public Const SHGFI_EXETYPE  As Long = &H2000
Public Const SHGFI_SYSICONINDEX  As Long = &H4000
Public Const SHGFI_SHELLICONSIZE  As Long = &H4
Public Const SHGFI_TYPENAME  As Long = &H400
Public Const SHGFI_LARGEICON  As Long = &H0
Public Const SHGFI_SMALLICON  As Long = &H1
Public Const ILD_TRANSPARENT As Long = &H1
Public Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE Or SH_USEFILEATTRIBUTES
Public FileInfo As typSHFILEINFO

Dim tempstr As String

Public Function geticonhandle(Filename As String, size As Long) 'Gets a handle to the icon
    geticonhandle = SHGetFileInfo(Filename, FILE_ATTRIBUTE_NORMAL, FileInfo, Len(FileInfo), Flags Or size)
End Function

Public Function drawfileicon(filetype As String, size As iconsize, destHDC As Long, x As Long, y As Long) 'Draws the icon int the destination.hdc
drawfileicon = ImageList_Draw(geticonhandle(filetype, size), FileInfo.iIcon, destHDC, x, y, ILD_TRANSPARENT)
End Function
