Attribute VB_Name = "modGenerate"
Option Explicit
Option Compare Text
'We need Option Compare Text for the sorting functions to work properly.

'  modGenerate.bas for Web Shelf 5
'  created by Andrew (aka BadHart)
'  Source submitted to Planet Source Code

'This module is responsible for the generation of the Web Shelves.

' **** CONSTANTS ****
Const DefTitle = "Web Shelf"
'the default title of a newly created web page.
Const WSURL = "http://badhart.tripod.com/webshelf"
'The URL of the program's web page.

' **** API CALLS ****
Private Declare Function IsCharAlphaNumeric Lib "user32" _
    Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
'Checks whether a single character is a letter or a number, or not.

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
    ByVal dwRop As Long) As Long
'Copies a picture between two controls, both having a device context (for
'example, pictureboxes or forms). dwRop, as Microsoft kindly and thoroughly
'explained (sarcasm), is the paint mode. All possible paint modes are
'RasterOp Constants.
'The difference between this and BitBlt is that the resulting picture can be
'stretched or squashed; while BitBlt takes rectangular sections of pictures.

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long
'This nice API call deletes files. Unfortunately I couldn't find the one that sends
'files to the Recycle Bin.

' **** TYPES ****
Type WSItem
    filename As String
    Size As Long
    Width As Integer
    Height As Integer
    DoLink As Boolean
End Type

' **** VARIABLES / COLLECTIONS ****
Private WSItems() As WSItem 'A dynamic array of the selected files. I used an
'array because collections were really p!ing me off. Pardon my French.
Private path As String, NewPath As String

Sub PlainAdvert(fn As Integer)
    'This sub was made in case the user doesn't want to display the
    'user information.
    
    Print #fn, "<TR><TD colspan=" & frmMain.sliWSImagesAcross.Value & ">"
    Print #fn, "<TABLE align='center' border=0 style='font-size: " & _
        frmMain.cboUserFontSize.Text;
    Print #fn, "; font-family: " & UserFontLine & "'>"
    Print #fn, "<TR><TD colspan=2 align='right'>"
    Print #fn, "Created with <b>" & "<a href='" & WSURL & "'>" _
        & App.Title & " " & App.Major & "</a></b> from BadSoft</TD></TR>"
    
    'Close the table...
    Print #fn, "</TABLE>"
    
    'Close the outer cell of the user information.
    Print #fn, "</TD></TR>"

End Sub

Sub SortByName(Ascending As Boolean)
    'This subroutine will sort all of the images int he WSItems array by their
    'filename. Remember that I used Option Compare Text, but it might not
    'be needed.
    
    'First we call NoSort to fill the array with the image files.
    NoSort
    
    'Then we do the sorting! This is called the 'bubble method' and I learned
    'it from Easy AMOS.
    Dim a As Integer, b As Integer, temp As WSItem
    
    For a = 1 To UBound(WSItems)
        For b = 1 To UBound(WSItems)
            'Cater for both ascending and descending orders.
            If (Ascending And WSItems(a).filename < WSItems(b).filename) Or _
                (Not Ascending And WSItems(a).filename > WSItems(b).filename) _
                Then
                temp = WSItems(b)
                WSItems(b) = WSItems(a)
                WSItems(a) = temp
            End If
        Next
    Next
End Sub

Sub SortBySize(Ascending As Boolean)
    'This subroutine will sort all of the images in the WSItems array by their size
    'in bytes.
    
    'First we call NoSort to fill the array with the image files.
    NoSort
    
    'Then we do the sorting!
    Dim a As Integer, b As Integer, temp As WSItem
    
    For a = 1 To UBound(WSItems)
        For b = 1 To UBound(WSItems)
            'Cater for both ascending and descending orders.
            If (Ascending And WSItems(a).Size < WSItems(b).Size) Or _
                (Not Ascending And WSItems(a).Size > WSItems(b).Size) Then
                temp = WSItems(b)
                WSItems(b) = WSItems(a)
                WSItems(a) = temp
            End If
        Next
    Next
End Sub

Function WriteWS(filename As String) As Boolean
    'Attempts to generate a completed Web Shelf, but returns False for no
    'success.
    
    On Error GoTo failed
    
beginning:
    'The topmost info, namely the title and any meta tags.
    Dim filenum As Integer
    filenum = FreeFile 'Assign a valid file number to the resulting file.
    
    'Open a normal format file for writing to.
    Open filename For Output As #filenum
    
    'In this section we write all the tags up to the </HEAD> tag. Of course
    'we'd use Write# rather than Print# as we need not concern ourselves
    'with retrieving data from HTML files. That would be stupid.
    
    'Update the status...
    With frmOptions
        .lblStatus = "Writing header"
        .asxProgress.Value = 0
        .asxProgress.Max = 1
    End With
    
    Print #filenum, "<HTML>"
    Print #filenum, "<HEAD>"
    Print #filenum, "<META name='Generator' content='Web Shelf 5'>"
    Print #filenum, "<TITLE>"
    'If the user has specified a title, make that the page title. Otherwise use
    'the default.
    Print #filenum, IIf(frmMain.txtWSTitle = Empty, DefTitle, _
    frmMain.txtWSTitle) & " created with " & App.Title & " " & App.Major
    Print #filenum, "</TITLE>"
    Print #filenum, "</HEAD>"
    Print #filenum, "<BODY>" 'believe it or not, I forgot this!
      
    'Indicate that the process has finished.
    frmOptions.asxProgress.Value = 1
    
top:
    'This is where we manage the <TABLE> tag and the Web Shelf caption.
    'Also, if necessary, we put the user information at the top. This means
    'we've got to find the number of images across.
    
    'First we get the image details by calling the requested sort method.
    With frmOptions
        .lblStatus = "Getting image information"
        .asxProgress.Value = 0
        .asxProgress.Max = ImageCount()
    End With

    If frmMain.optNoImageSort Then
        'No sort order. Note that this is called by the other two sort functions,
        'as this provides the list of images.
        NoSort
    ElseIf frmMain.optSortByName Then
        'Sort by filename.
        SortByName frmMain.optOAscending
    Else
        'Sort by file size.
        SortBySize frmMain.optOAscending
    End If
    
    'Make thumbnails only if necessary.
    If frmMain.chkThumbs Then
        MakeThumbs Val(frmMain.txtImgToWidth), Val(frmMain.txtImgToHeight)
    End If
    
    'Next we start the table.
    Print #filenum, "<TABLE align='center' width='90%'";
    'Why 90%? Because we want the Web Shelf to fill the page, and also to
    'try and make all of the cells even widths. Options to change this may be
    'made available in the next release.
    
    With frmMain
        'Border size, cell padding and spacing.
        Print #filenum, " border=" & .sliBorderSize.Value;
        Print #filenum, " cellpadding=" & .sliWSCellPadding;
        Print #filenum, " cellspacing=" & .sliWSCellSpacing
        
        'Style. Web Shelf 5 uses style attributes, so this is a dangerous
        'assumption that the user's browser supports CSS. These style
        'tags are much more resourceful and result in smaller files, but next
        'release you'll have the option of doing it the old way.
        Print #filenum, " style='" & IIf(UserFontLine <> Empty, "font-family: " & _
            UserFontLine & "; ", "") & "font-size: " & frmMain.cboUserFontSize & "'>";
            
        'I should also mention that some of the string expressions have semi-
        'colons at the end. This is so that the next printed string appears on the
        'same line. So that....
        
            'Print #filenum, "Hello";
            'Print #filenum, " you"
            
        'reads in Notepad...
            
            'Hello you
            
        'Okay?
        
    End With
    'The <TABLE> tag is closed (see above).
    
    'If there is a given title for the Web Shelf, we put that on top.
    If frmMain.txtWSTitle <> Empty Then
        Print #filenum, "<CAPTION align='top'";
        'Get font information for the Web Shelf title...
        Print #filenum, " style='" & IIf(TitleFontLine <> Empty, _
            "font-family: " & TitleFontLine, "") & _
            "; font-size: " & frmMain.cboTitleFontSize & "'>"
        'The title itself.
        Print #filenum, frmMain.txtWSTitle
        'Close the <CAPTION> tag.
        Print #filenum, "</CAPTION>"
    End If
    
    'If we are to put the user information at the top of the Web Shelf, do so
    'now.
    If frmMain.optUserAlignTop Then UserInfo filenum
    
    'Now to put the main part of the Web Shelf in. Let's go!
    MakeTable filenum, frmMain.sliWSImagesAcross.Value, _
        frmMain.lblImagesHigh
        
    'If the user information is to go at the bottom, put it there!
    If frmMain.optUserAlignBottom Then
        UserInfo filenum
    Else
        'Otherwise we put an advert in. Did you think the user would get off that
        'easily?
        PlainAdvert filenum
    End If
    
    'Finally we close the table and the page.
    Print #filenum, "</TABLE>"
    Print #filenum, "</BODY>"
    Print #filenum, "</HTML>"
    
    'If we've made it to here without any errors, skip the next label. Otherwise
    'we get a very unuseful error 0.
    GoTo closing

failed:
    'Something went wrong here, and so we must report the error.
    MsgBox Err.Description & ".", vbExclamation
    Close #filenum
    WriteWS = False
    Exit Function
    
closing:
    'Now that we've done everything, we must close the file.
    Close #filenum
    WriteWS = True

End Function

Sub UserInfo(fn As Integer)
    'This part deals with user information. That is, if the user wants it
    'displayed.
    'This just involves creating a cell that spans the whole Web Shelf, inserting
    'an invisible table inside it, and then showing the criteria.

    With frmOptions
        .lblStatus = "Writing user information"
        .asxProgress.Value = 0
        .asxProgress.Max = 1
    End With

    Print #fn, "<!-- Beginning of user information -->"
    'Another comment
    
    'This is the start of the large cell.
    Print #fn, "<TR><TD colspan=" & frmMain.sliWSImagesAcross.Value & ">"
    
    'Start of the inside invisible table.
    Print #fn, "<TABLE align='center' border=0 style='font-size: " & _
        frmMain.cboUserFontSize.Text
    Print #fn, "; font-family: " & UserFontLine & "'>"
    
    'Only print the user information if the user wants it displayed.
    If frmMain.optNoUser = False Then
        'A header to let everybody know...
        Print #fn, _
            "<TR><TD><FONT size='3'><B>Information</B></FONT></TD></TR>"
    
        'Name
        If frmMain.txtUserName <> Empty Then
            Print #fn, "<TR><TD><B>Name<B></TD>"
            Print #fn, "<TD>" & frmMain.txtUserName & "</TD></TR>"
        End If
        
        'Image copyright
        If frmMain.txtUserCopyright <> Empty Then
            Print #fn, "<TR><TD><B>Image Copyright<B></TD>"
            Print #fn, "<TD>" & frmMain.txtUserCopyright & "</TD></TR>"
        End If
        
        'Comments
        If frmMain.txtUserComments <> Empty Then
            Print #fn, "<TR><TD><B>Comments<B></TD>"
            Print #fn, "<TD>" & frmMain.txtUserComments & "</TD></TR>"
        End If
        
        'If necessary, put in the date.
        If frmMain.chkUserIncludeDate Then
            Print #fn, "<TR><TD></TD>"
            Print #fn, "<TD>created on " & Date & "</TD></TR>"
        End If
        
    End If
    
    'The mandatory advertisement, which can still be deleted but is
    'automatically generated at my convenience.
    Print #fn, "<TR><TD colspan=2 align='right'>"
    Print #fn, "Created with <B>" & "<A href='" & WSURL & "'>" _
        & App.Title & " " & App.Major & "</A></B> from BadSoft</TD></TR>"
    
    'Close the invisible table...
    Print #fn, "</TABLE>"
    
    'Close the large cell.
    Print #fn, "</TD></TR>"
    Print #fn, "<!-- End of user information -->"
    
    frmOptions.asxProgress.Value = 1
End Sub

Sub DoGenerate()
    'Starts the generation of the web page and handles success and failure.
    'But before any of that, we have to set the path.
    
    path = frmMain.txtPath
    
    'Do we create in the original folder?
    With frmOptions
        If .optOldDir Then
            'Yes we do...
            NewPath = path
        Else
            'We create a new directory.
            NewPath = frmOptions.txtPath & _
                IIf(frmOptions.txtNewFolder <> Empty, "\" & _
                    frmOptions.txtNewFolder, "")
                
            'If this folder doesn't exist then create a new one.
            If Dir(NewPath, vbDirectory) = Empty Then
                MkDir NewPath 'Remember this? This is just like a DOS command.
            End If
        End If
        
        'Make any new directories.
        .lblStatus = "Making new directories..."
        MakeDirs NewPath
            
        'Only if we're dealing with a different folder, copy the original images.
        If .optNewDir Then
            .lblStatus = "Copying image files to new folder..."
            CopyFiles path, NewPath
        End If
        
    End With
    
    On Error GoTo whoops
    
    'We run the command to make the Web Shelf, and if there are any
    'problems with the code then I should know.
    frmOptions.lblStatus = "Generating Web Shelf..."

    If WriteWS(NewPath & "\" & frmOptions.txtFilename & ".html") = False Then
        MsgBox "Unable to save the generated Web Shelf."
    End If
    
whoops:
End Sub

Function UserFontLine() As String
    'Returns the fonts that will be used in the user table. If both are set to the
    'default this string is empty.
    
    Dim First As Boolean, Second As Boolean
    With frmMain
        First = (.lblUserFont(0) <> DefaultFont)
        Second = (.lblUserFont(1) <> DefaultFont)
        
        'Now to put the string together.
        UserFontLine = IIf(First, .lblUserFont(0), "") & IIf(First And Second, ", ", "") & _
            IIf(Second, .lblUserFont(1), "")
    End With
    
End Function

Function TitleFontLine() As String
    'Returns the fonts that will be used in the Web Shelf's title. If both are set
    'to the default this string is empty.
    
    Dim First As Boolean, Second As Boolean
    With frmMain
        First = (.lblTitleFont(0) <> DefaultFont)
        Second = (.lblTitleFont(1) <> DefaultFont)
        
        'Now to put the string together.
        TitleFontLine = IIf(First, .lblTitleFont(0), "") & IIf(First And Second, ", ", "") & _
            IIf(Second, .lblTitleFont(1), "")
    End With
    
End Function
Sub NoSort()
    'Now that we have a collection of images, there are options to sort them
    'by filename or file size.
    'This subroutine will take each image as they come, ie. as specified by
    'the Dir function.
    
    Dim oneWSItem As WSItem, myDir As String, a As Integer, tp As Long
       
    'To save coding space, we take a trip down memory lane and go back
    'to using GoSub statments. Remmeber those?
    a = 1: ReDim WSItems(1 To ImageCount)
       
    myDir = Dir(frmMain.txtPath & "\" & "*.jpg"): GoSub find
    myDir = Dir(frmMain.txtPath & "\" & "*.gif"): GoSub find
    myDir = Dir(frmMain.txtPath & "\" & "*.bmp"): GoSub find
    
    Exit Sub
    
find:
    While myDir <> Empty
        'Just add each item to the list.
        With oneWSItem
            .filename = myDir
            .Size = FileLen(frmMain.txtPath & "\" & myDir)
            
            'To find the image's dimensions, we have to load it into an invisible
            'image control.
            frmMain.imgDimension.Picture = _
                LoadPicture(frmMain.txtPath & "\" & myDir)
            .Width = frmMain.imgDimension.Width
            .Height = frmMain.imgDimension.Height
        End With
    
        WSItems(a) = oneWSItem
        myDir = Dir: a = a + 1
        frmOptions.asxProgress.Value = a
    Wend
    
    Return
End Sub

Sub MakeTable(fn As Integer, Width As Integer, Height As Integer)
    'This is the backbone of the program; it creates the actual table of
    'images. But it's going to be very complicated!
    'Here we go...
    
    Dim cw As Integer, ch As Integer, Cellnumber As Integer
    Dim WSItemRef As WSItem, temp As String, nf As String
    
    'The table must have been opened before this subroutine was called, to
    'cater for user information that might have been requested at the top
    'of the Web Shelf. So we go straight to rows!
    
    With frmOptions
        .lblStatus = "Generating Web Shelf code"
        .asxProgress.Value = 0
        .asxProgress.Max = ImageCount()
    End With

    Print #fn, "<!-- Beginning of Web Shelf code -->"
    'A comment for anyone who might want to take a peek at the
    'generated code.
    
    For ch = 1 To Height
        'Start a new row...
        Print #fn, "<TR>"
        
        For cw = 1 To Width
            'We also need to do a small check. If we've run out of images, we
            'just leave a blank cell.
            Cellnumber = cw + (ch - 1) * Width
            
            If Cellnumber <= UBound(WSItems) Then
                'We haven't run out of images, so start a new cell.
                Print #fn, "<TD align='center'>"
                
                'The easiest way to make everything nice and neat is to put
                'everything into the cell separated by line breaks. This might lead
                'to a problem I can't get rid of, where a line of white space
                'separates the image from any information.
                          
                temp = SizeLine(Cellnumber)
                'This was done to check if the current image will be resized.
                
                'Do we link to the image's full size if it has been resized?
                If frmMain.chkLinktoFullSize And WSItems(Cellnumber).DoLink Then
                    'Start the link. The only way to do this is with the anchor (<A>)
                    'tag.
                    Print #fn, "<A href='" & IIf(frmOptions.optNewDir, "full/", "") & _
                        WSItems(Cellnumber).filename & "'>"
                End If
                
                'The source of the images depends on whether or not we have
                'created a new directory. The full sized images in a new directory
                'will be in the subdirectory 'full'.
                'But now we also have to worry about present thumbnails.
                nf = WSItems(Cellnumber).filename
                nf = IIf(frmOptions.chkRename, IFriendly(nf), nf)
                If frmMain.chkThumbs Then
                    'If we want thumbnails, we always point to the thumbnailed
                    'image. But be warned, we also have to make this file type a
                    'bitmap.
                    temp = "thumbs/" & Left(nf, Len(nf) - 3) & "bmp"
                ElseIf frmOptions.optNewDir Then
                    'If we don't have thumbnails but have all images in a new
                    'directory then we link to the full directory.
                    temp = "full/" & nf
                Else
                    'Otherwise we link to the original images, which is where the
                    'Web Shelf will and should be.
                    temp = nf
                End If
                Print #fn, "<IMG src='" & temp;
                
                'If images are to be linked to then a border is placed around
                'them. This is purely for visual clarity, and we also place some ALT
                'text in the <IMAGE> tag as well.
                Print #fn, IIf(WSItems(Cellnumber).DoLink And _
                    frmMain.chkLinktoFullSize, _
                    "' alt='Click here to see the full size picture!' border=1", _
                    "' border=0");
                    
                'Do we want to resize this image? (Don't do it if we are dealing
                'with thumbnails!)
                If frmMain.optResize And Not frmMain.chkThumbs Then
                    'Get the width and height settings.
                    Print #fn, SizeLine(Cellnumber);
                End If

                'Close the image tag.
                Print #fn, ">"
            
                'If we have linked, close the anchor tag.
                If frmMain.chkLinktoFullSize And WSItems(Cellnumber).DoLink Then
                    Print #fn, "</A>"
                End If
                
                'Get any image information to be displayed.
                ImageInfo fn, Cellnumber
                              
                'Close the cell...
                Print #fn, "</TD>"
                
                frmOptions.asxProgress.Value = Cellnumber
            End If
        Next
        
        'End the row.
        Print #fn, "</TR>"
    Next
    
    Print #fn, "<!-- End of Web Shelf code -->"
    'Signal the end of the Web Shelf table.
End Sub

Function SizeLine(inum As Integer) As String
    'This figures out whether or not we should resize an image. It returns
    'the width= and height= parameters of the img tag.
    
    With frmMain
        If .chkLimitResize Then
            'Only perform these checks if instructed.
            
            If .optNoResize1 And (WSItems(inum).Width > Val(.txtImgToWidth) _
                Or WSItems(inum).Height > Val(.txtImgToHeight)) Then
                'Resize if either of the dimensions are larger than the
                'specified.
                
                SizeLine = " width=" & Val(.txtImgToWidth) & " height=" & _
                    Val(.txtImgToHeight)
                WSItems(inum).DoLink = True

            End If
            
            If .optNoResize2 And (WSItems(inum).Width > Val(.txtImgToWidth) _
                And WSItems(inum).Height > Val(.txtImgToHeight)) Then
                'Resize if BOTH dimensions are larger than the specified.
                
                SizeLine = " width=" & Val(.txtImgToWidth) & " height=" & _
                    Val(.txtImgToHeight)
                WSItems(inum).DoLink = True
            
            End If
        Else
            'However, if we have to resize all the images anyway, do so.
            SizeLine = " width=" & Val(.txtImgToWidth) & " height=" & _
                Val(.txtImgToHeight)
            WSItems(inum).DoLink = True
        End If
    End With
End Function

Sub CopyFiles(ByVal oPath As String, ByVal nPath As String)
    'This subroutine copies all of the image files to the new directory.
    
    Dim myDir As String
    myDir = Dir(oPath & "\*.gif", vbNormal): GoSub copy
    myDir = Dir(oPath & "\*.jpg", vbNormal): GoSub copy
    myDir = Dir(oPath & "\*.bmp", vbNormal): GoSub copy
    Exit Sub
    
copy:
    'This subroutine does all of the copying.
    While myDir <> Empty
        FileCopy oPath & "\" & myDir, nPath & "\full\" & _
            IIf(frmOptions.chkRename, IFriendly(myDir), myDir)
        myDir = Dir
    Wend
    Return
    
End Sub

Sub MakeDirs(path As String)
    'This makes the directories called 'full' and - only if necessary - 'thumbs'.
    'Before we can do this however, we need to check if they already exist.
    'If they do, we just erase all the files in them.
    
    'Do the check...
    Dim myDir As String, Folder As String
    
    'But suppose the user wants thumbnails but wants to create in the
    'original directory. In that case we wouldn't need to make a 'full'
    'directory.
    
    If frmOptions.optNewDir Then
        Folder = "full": GoSub delete
        'Only if we require a replacement directory.
    End If
    
    If frmMain.chkThumbs Then
        Folder = "thumbs": GoSub delete
        'Create a thumbs directory if necessary.
    End If
    
    Exit Sub
    
delete:
    myDir = Dir(path & "\" & Folder, vbDirectory)
    If myDir <> Empty Then
        'Erase all of the files! Note that this does not send them to the Recycle
        'Bin; you'll need to use another API function for that.
        DeleteFile path & "\" & Folder & "\*.*"
    Else
        'We create the directory instead.
        MkDir path & "\" & Folder
    End If
    Return
    
End Sub

Sub ImageInfo(fn As Integer, inum As Integer)
    'This sub puts all of the required image information into the open small
    'image table.
    
    Dim br As Boolean 'this variable is true if there was a previous item and
    'controls breaks.
    

    With frmMain
        'Do a check for no image information required at all.
        If (.chkShowDimensions Or .chkShowFileName Or .chkShowFileSize Or _
             .chkShowFileType) Then
            'Add a line break if there is at least one item.
            Print #fn, "<br>"
        End If
        
        'First the filename.
        If .chkShowFileName Then
            Print #fn, "<b>" & _
                Left(WSItems(inum).filename, Len(WSItems(inum).filename) - 4) & "</b>"
                br = True
        End If
        
        'File size.
        If .chkShowFileSize Then
            If br Then Print #fn, "<br>"
            Print #fn, Trim(Str(WSItems(inum).Size)); " bytes"
            br = True
        End If
        
        'File types.
        If .chkShowFileType Then
            If br Then Print #fn, "<br>"
            Select Case LCase(Right(WSItems(inum).filename, 3))
            Case "bmp"
                Print #fn, "Bitmap"
            Case "gif"
                Print #fn, "GIF image"
            Case "jpg"
                Print #fn, "JPEG image"
            End Select
            br = True
        End If
        
        'The image's dimensions.
        If .chkShowDimensions Then
            If br Then Print #fn, "<br>"
            Print #fn, WSItems(inum).Width & " x " & _
                WSItems(inum).Height & " pixels"
            br = True
        End If
        
    End With
End Sub

Function IFriendly(what As String) As String
    'This function will remove the funny characters from the string 'what', and
    'also replaces spaces with underscores. This is for Web Shelves intended
    'for the Web.
    
    Dim a As Integer, t As String
    For a = 1 To Len(what)
        t = Mid(what, a, 1)
        If t = " " Or t = "_" Then
            'If the character is a space or underscore, replace it with an
            'underscore.
            IFriendly = IFriendly & "_"
        ElseIf IsCharAlphaNumeric(Asc(t)) Then
            'The character is a letter or number.
            IFriendly = IFriendly & LCase(t)
        ElseIf t = "." Then
            'Periods are welcome too (but not for women - bad joke)
            IFriendly = IFriendly & "."
        Else
            'This character is invalid, so we don't add it to the new string.
        End If
    Next
End Function

Sub MakeThumbs(Width As Integer, Height As Integer)
    'Here we go. This is the second-to-last subroutine to write before the
    'program is finished, and it involves recalling some code I had written in
    'Planet Source Code. Using only my technique, I will now attempt to
    'create thumbnails of the images specified, and with the given size.
    
    With frmOptions
        'Make sure that the Scalemode property is PIXELS. You don't know how
        'many people were on my case when I forgot this.
        .ScaleMode = vbPixels
        'Resize the picDest control to fit the thumbnail dimensions.
        .picDest.Width = Width
        .picDest.Height = Height
        'Make picDest and the form use a persistent bitmap (so that we don't
        'lose the generated thumbnail).
        .picDest.AutoRedraw = True
        .AutoRedraw = True
        'picSource should resize to fit the full scale picture.
        .picSource.AutoSize = True
        'Both should be made invisible and without borders.
        .picDest.Visible = False
        .picSource.Visible = False
        .picDest.BorderStyle = vbBSNone
        .picSource.BorderStyle = vbBSNone

        'Now to start!
        
        'Do some status bar business...
        .lblStatus = "Making thumbnails"
        .asxProgress.Max = ImageCount()
        .asxProgress.Value = 0
        
        Dim a As Integer, temp As String, b As Long
        For a = 1 To ImageCount()
            'Load the full size picture.
            .picSource = LoadPicture(frmMain.txtPath & "\" & _
                WSItems(a).filename)
            'Copy the picture to the picDest control, but resize it in the process.
            'Compare the two given methods and figure out which is faster and
            'more professional. I am just glad I remembered the API call in time!
            
            'PaintPicture method
            '.picDest.PaintPicture .picSource.Picture, 0, 0, Width, Height, 0, 0, _
                .picSource.ScaleWidth , .picSource.ScaleHeight
                
            'StretchBlt API method
            b = StretchBlt(.picDest.hdc, 0, 0, Width, Height, .picSource.hdc, 0, 0, _
                .picSource.ScaleWidth, .picSource.ScaleHeight, vbSrcCopy)
            'if b=0 then the thumbnail couldn't be created.
            
            'Save the picture in the thumbs directory, not forgetting to rename
            'the image if necessary.
            'At present Visual Basic only saves images in bitmap format; but
            'there are methods and special ActiveX controls that can save to
            'GIF and JPEG formats. Web Shelf 6 improvement, methinks!
            temp = IIf(frmOptions.chkRename, IFriendly(WSItems(a).filename), _
                WSItems(a).filename)
            SavePicture .picDest.Image, NewPath & "\thumbs\" & _
                Left(temp, Len(temp) - 3) & "bmp"
                
            'Increase the status bar.
            .asxProgress.Value = a
        Next
    End With
End Sub
