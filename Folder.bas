Attribute VB_Name = "FolderControlSubs"
Option Explicit

Private Declare Function OleInitialize Lib "ole32.dll" (lp As Any) As Long
Private Declare Sub OleUninitialize Lib "ole32.dll" ()
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
        (PicDes As PicDesc, RefIID As IID, ByVal fPictureOwnsHandle As Long, _
        IPic As IPicture) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
 
Private Const MAX_PATH = 260
Private Const cMagic As Double = -2.51702880262616E-101
Private Const SHGFI_ICON = &H100                         '  get icon
Private Const SHGFI_SMALLICON = &H1
Public Const PICTYPE_ICON = 3

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    BData1 As Byte
    BData2 As Byte
    BData3 As Byte
    BData4 As Byte
    BData5 As Byte
    BData6 As Byte
    BData7 As Byte
    BData8 As Byte
End Type

Private Type SHFILEINFO
    hIcon As Long                      '  out: icon handle
    iIcon As Long          '  out: icon index within system image list
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80         '  out: type name
End Type

Private Type PicDesc
  cbSizeOfStruct As Long
  picType As Long
  hGdiObj As Long
  hPalOrXYExt As Long
End Type

Dim picData As PicDesc
Dim IconInfo As SHFILEINFO
Dim id As IID


'---------------------------------------------------------------------------

Public Function GetIcon(path As String) As IPictureDisp
Dim StdPic As IPictureDisp
'get file associated icon info
SHGetFileInfo path, 0, IconInfo, Len(IconInfo), SHGFI_ICON Or SHGFI_SMALLICON

If IconInfo.hIcon = 0 Then Exit Function

picData.cbSizeOfStruct = Len(picData)
picData.picType = PICTYPE_ICON
picData.hGdiObj = IconInfo.hIcon
' can be null picData.hPalOrXYExt = 0&

id.Data1 = &H7BF80981
id.Data2 = &HBF32
id.Data3 = &H101A
id.BData1 = &H8B
id.BData2 = &HBB
id.BData3 = &H0
id.BData4 = &HAA
id.BData5 = &H0
id.BData6 = &H30
id.BData7 = &HC
id.BData8 = &HAB

OleInitialize ByVal 0&
OleCreatePictureIndirect picData, id, True, StdPic
OleUninitialize
Set GetIcon = StdPic
End Function

Public Function DirExist(path As String) As Boolean
On Error GoTo err
If Dir(path, vbDirectory) <> "" Then
DirExist = True
End If
Exit Function
err:
DirExist = False
End Function

Public Function FileExist(strPath As String) As Boolean
Dim FileLength As Long
On Error Resume Next
FileLength = FileLen(strPath)
FileExist = (err.Number = 0)
End Function

