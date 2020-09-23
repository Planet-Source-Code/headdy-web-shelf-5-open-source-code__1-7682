Attribute VB_Name = "modMisc"
Option Explicit

'  modMisc.bas for Web Shelf 5
'  created by Andrew (aka BadHart)
'  Source submitted to Planet Source Code

'To keep everything neat I like to do this with any declared variables at the
'top of the module. Wouldn't it be great if someone were to write a
'program that automated this process?

' **** TYPES ****
Private Type BrowseInfo 'Needed for the BrowseForFolder function.
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'**** CONSTANTS ****
Private Const BIF_RETURNONLYFSDIRS = &H1
'For finding a folder.

Private Const BIF_DONTGOBELOWDOMAIN = &H2
'For starting Find Computer.

Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8

Private Const BIF_BROWSEFORCOMPUTER = &H1000
'Browsing for Computers.

Private Const BIF_BROWSEFORPRINTER = &H2000
'Browsing for Printers.

Private Const BIF_BROWSEINCLUDEFILES = &H4000
'Browsing for Everything!

Private Const MAX_PATH = 260
'Most probably the maximum length of the returned path.

Public Const DefaultFont = "Default browser font"
'The string to use for indicating the default font. This is mine!

' **** API CALLS ****
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal _
lpString1 As String, ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As _
BrowseInfo) As Long
'The API call to browse for a folder.

Private Declare Function SHGetPathFromIDList Lib "shell32" _
(ByVal pidList As Long, ByVal lpBuffer As String) As Long

' **** PRIVATE VARIABLES ****
Dim MyPath As String
Dim Filetype(1 To 3) As Integer    'used for counting the number of image files
Dim TotalFileBytes As Long          'total size of all the images, in bytes.

' **** PUBLIC VARIABLES ****
Public Changed As Boolean
'If the Web Shelf's parameters have changed.

Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) _
As String

    'Opens the system dialog for browsing for a folder. I can't quite remember
    'who submitted this code, but thanks anyway!
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    
    With udtBI
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    
    BrowseForFolder = sPath
    
End Function

Sub ListImages(path As String)
    'This subroutine generates the statistics for the lsvImgStats control.
       
    'First we have to clear the list!
    With frmMain
        .lsvImgStats.ListItems.Clear
        Dim a As Integer
        'Empty the Filetype array of any existing 'files'.
        For a = 1 To 3
            Filetype(a) = 0
        Next
        'There are no files, so the total file size is 0.
        TotalFileBytes = 0
    
        'Now to list all of the images. This could have been done in another
        'way, as demonstrated somewhere else, but this was the first method
        'that came to mind. Also in this version of Web Shelf we only look for
        'one kind of extension.
        
        'First we look for bitmaps.
        MyPath = Dir(path & "\*.bmp", vbNormal)
        While MyPath <> Empty
            'If there is a next bitmap, we get its stats.
            Filetype(1) = Filetype(1) + 1 'increase bitmaps by one
            'add its file size
            TotalFileBytes = TotalFileBytes + FileLen(path & "\" & MyPath)
            'then find the next existing bitmap.
            MyPath = Dir
        Wend
        
        'Now GIFs...
        MyPath = Dir(path & "\*.gif", vbNormal)
        While MyPath <> Empty
            'Same thing here.
            'Add the image to the list...
            Filetype(2) = Filetype(2) + 1
            TotalFileBytes = TotalFileBytes + FileLen(path & "\" & MyPath)
            MyPath = Dir
        Wend
        
        'Now JPEGs.
        MyPath = Dir(path & "\*.jpg", vbNormal)
        While MyPath <> Empty
            Filetype(3) = Filetype(3) + 1
            TotalFileBytes = TotalFileBytes + FileLen(path & "\" & MyPath)
            MyPath = Dir
        Wend
    End With
      
    'Time to put all of this information into the listview control. If there's an
    'easier way to do this then please let me know.
    With frmMain.lsvImgStats.ListItems
        .Add , , "Total number of files"
        .Item(1).ListSubItems.Add , , Filetype(1) + Filetype(2) + Filetype(3)
        .Item(1).Bold = True
        .Item(1).ListSubItems(1).Bold = True
        .Add , , "Total file size"
        .Item(2).ListSubItems.Add , , TotalFileBytes & " bytes"
        .Add , , "Number of GIFs"
        .Item(3).ListSubItems.Add , , Filetype(2)
        .Add , , "Number of JPEGs"
        .Item(4).ListSubItems.Add , , Filetype(3)
        .Add , , "Number of bitmaps"
        .Item(5).ListSubItems.Add , , Filetype(1)
        .Add , , "Download at 14.4 Kb/s"
        .Item(6).ListSubItems.Add , , Format(TotalFileBytes / 1024 / 14.4, "0.##") & " seconds"
        .Add , , "Download at 28.8 Kb/s"
        .Item(7).ListSubItems.Add , , Format(TotalFileBytes / 1024 / 28.8, "0.##") & " seconds"
        .Add , , "Download at 56 Kb/s"
        .Item(8).ListSubItems.Add , , Format(TotalFileBytes / 1024 / 56, "0.##") & " seconds"
    End With
    
    'Get rid of the warning if there are any images.
    With frmMain
        .lblNoImages.Visible = NoImages
        
        'Also enable the slider if necessary.
        With .sliWSImagesAcross
            .Enabled = Not NoImages
            'Set a default size for images across. This should automatically arouse
            'calculation of the height.
            .Max = Limit(ImageCount(), 2, 12)
            .Value = .Max
        End With
    End With
End Sub

Function NoImages() As Boolean
    'Returns true if there are no images in the current (chosen) directory.
    NoImages = (Filetype(1) + Filetype(2) + Filetype(3) = 0)
End Function

Function ImageCount() As Integer
    'Pretty self explanatory...
    ImageCount = Filetype(1) + Filetype(2) + Filetype(3)
End Function

Sub Main()
    'This is the startup procedure. The main form is loaded before it is shown.
    Load frmMain
    frmMain.Show
    'Set the tabs control to display the first tab (images).
    frmMain.sstTabs.Tab = 0
End Sub

Function Limit(Value, Low, High)
    'This may work for strings too! I haven't tried it.
    If Low > High Then
        Err.Raise 8001, , "Low parameter in Limit() function is greater than High."
    End If
    If Value < Low Then Value = Low
    If Value > High Then Value = High
    Limit = Value
End Function

Sub NewSettings()
    'This creates a new Web Shelf settings file; in effect, it resets everything.
    Dim a As Integer
    
    'Make sure there are no images.
    For a = 1 To 3
        Filetype(a) = 0
    Next
        
    With frmMain
        'IMAGES TAB
        .txtPath = Empty
        'Don't forget to clear the listview control.
        .lsvImgStats.ListItems.Clear
        
        'DIMENSIONS TAB
        .lblNoImages.Visible = True
        .sliWSImagesAcross.Enabled = False
        .sliWSImagesAcross = 6
        .lblImagesHigh = "?"
        
        'APPEARANCE TAB
        .txtWSTitle = "A Collection of Images"
        .sliBorderSize = 4
        .sliWSCellPadding = 2
        .sliWSCellSpacing = 1
        
        'FONTS TAB
        For a = 0 To 1
            .lblUserFont(a) = DefaultFont
            .lblTitleFont(a) = DefaultFont
        Next
        .cmdChooseFont(1).Enabled = False
        .cmdChooseTitleFont(1).Enabled = False
        .lblTitleFont(1).Enabled = False
        .lblUserFont(1).Enabled = False
        .cboTitleFontSize = "18pt"
        .cboUserFontSize = "10pt"
        
        'RESIZING TAB
        .optNoResize = True
        .chkLimitResize = 1
        .chkLinktoFullSize = 1
        .optNoResize1 = True
        .chkThumbs = 0
        .txtImgToHeight = 60
        .txtImgToWidth = 80
        .Container2.Visible = True
        .Container1.Visible = True
        
        'TEXT TAB
        .chkShowDimensions = 0
        .chkShowFileName = 1
        .chkShowFileSize = 0
        .chkShowFileType = 0
        .optNoImageSort = True
        .optOAscending = True
        .asxOContainer.Visible = False
        
        'USER INFO TAB
        .txtUserName = Empty
        .txtUserCopyright = Empty
        .txtUserComments = Empty
        .chkUserIncludeDate = 0
        .optNoUser = True
        .fraUserOpts.Visible = False
        
        Changed = False
    End With
End Sub

Function SaveSettings(filename As String) As String
    'This saves all of the details in the frmMain form to a file. If there is an error,
    'the file is closed and erased and the error description is returned.
    
    On Error GoTo whoops
    Dim a As Integer
    a = FreeFile
    
    Open filename For Output As #a
    'First we save the version number.
    Write #a, App.Major & "." & App.Minor & App.Revision
    
    With frmMain
        'IMAGES TAB
        'Only save the path from this tab.
        Write #a, .txtPath
        
        'DIMENSIONS TAB
        Write #a, .sliWSImagesAcross
        
        'APPEARANCE TAB
        Write #a, .txtWSTitle
        Write #a, .sliBorderSize
        Write #a, .sliWSCellPadding
        Write #a, .sliWSCellSpacing
        
        'FONTS TAB
        Write #a, .lblUserFont(0)
        Write #a, .lblUserFont(1)
        Write #a, .lblTitleFont(0)
        Write #a, .lblTitleFont(1)
        Write #a, .cboUserFontSize
        Write #a, .cboTitleFontSize
        
        'RESIZING TAB
        Write #a, .optNoResize
        Write #a, .optResize
        Write #a, .txtImgToWidth
        Write #a, .txtImgToHeight
        Write #a, .chkLimitResize
        Write #a, .optNoResize1
        Write #a, .optNoResize2
        Write #a, .chkThumbs
        Write #a, .chkLinktoFullSize
        
        'TEXT TAB
        Write #a, .chkShowFileName
        Write #a, .chkShowFileSize
        Write #a, .chkShowFileType
        Write #a, .chkShowDimensions
        Write #a, .optNoImageSort
        Write #a, .optSortByName
        Write #a, .optSortBySize
        Write #a, .optOAscending
        Write #a, .optODescending
        
        'Don't forget, user info is handled separately.
    End With
    
whoops:
    Close #a
    If Err.Number <> 0 Then Kill filename
    SaveSettings = Err.Description
End Function

Function OpenSettings(filename As String) As String
    'Opens previously saved information, so it's important we get everything
    'in the order that it was saved.
    
    On Error GoTo whoops
    Dim a As Integer, temp As String
    a = FreeFile
    
    Open filename For Input As #a
    'Get the version number.
    Input #a, temp
    'Check if this is a file created by a newer version.
    If Val(temp) > Val(App.Major & "." & App.Minor & App.Revision) Then
        MsgBox "This file was created with a newer version of Web Shelf and can't be opened.", vbExclamation
        frmMain.Status.SimpleText = "Web Shelf 5 - Ready"
        Exit Function
    End If
    
    With frmMain
        'IMAGES TAB
        Input #a, temp: .txtPath = temp
        ListImages temp
        
        'DIMENSIONS TAB
        Input #a, temp: .sliWSImagesAcross = Val(temp)
        
        'APPEARANCE TAB
        Input #a, temp: .txtWSTitle = temp
        Input #a, temp: .sliBorderSize = Val(temp)
        Input #a, temp: .sliWSCellPadding = Val(temp)
        Input #a, temp: .sliWSCellSpacing = Val(temp)
        
        'FONTS TAB
        Input #a, temp: .lblUserFont(0) = temp
        .lblUserFont(1).Enabled = (temp <> DefaultFont)
        .cmdChooseFont(1).Enabled = (temp <> DefaultFont)
        Input #a, temp: .lblUserFont(1) = temp
        Input #a, temp: .lblTitleFont(0) = temp
        .lblTitleFont(1).Enabled = (temp <> DefaultFont)
        .cmdChooseTitleFont(1).Enabled = (temp <> DefaultFont)
        Input #a, temp: .lblTitleFont(1) = temp
        Input #a, temp: .cboUserFontSize = temp
        Input #a, temp: .cboTitleFontSize = temp
       
        'RESIZING TAB
        Input #a, temp: .optNoResize = temp
        Input #a, temp: .optResize = temp
        Input #a, temp: .txtImgToWidth = temp
        Input #a, temp: .txtImgToHeight = temp
        Input #a, temp: .chkLimitResize = Val(temp)
        Input #a, temp: .optNoResize1 = temp
        Input #a, temp: .optNoResize2 = temp
        Input #a, temp: .chkThumbs = Val(temp)
        Input #a, temp: .chkLinktoFullSize = Val(temp)
        
        'TEXT TAB
        Input #a, temp: .chkShowFileName = Val(temp)
        Input #a, temp: .chkShowFileSize = Val(temp)
        Input #a, temp: .chkShowFileType = Val(temp)
        Input #a, temp: .chkShowDimensions = Val(temp)
        Input #a, temp: .optNoImageSort = temp
        Input #a, temp: .optSortByName = temp
        Input #a, temp: .optSortBySize = temp
        Input #a, temp: .optOAscending = temp
        Input #a, temp: .optODescending = temp
    End With
    
whoops:
    Close #a
    OpenSettings = Err.Description
End Function

