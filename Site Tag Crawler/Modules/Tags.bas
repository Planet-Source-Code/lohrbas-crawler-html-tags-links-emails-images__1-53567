Attribute VB_Name = "Tags"
'_____________________________________________________________
'                        Coded By: Lohrbas                    |
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|
' I just made this program cause I needed it and I've been    |
' looking for something like this (that's good and works)     |
' but I could never find anything efficient to my needs. So   |
' I decided to make this. With that said, I don't care if you |
' add this to your code, take off this information or modify  |
' anything. Hopefully this will make your lives easier and    |
' save 10 hours of your lives. I made this code really        |
' flexiable even though I only planned to use it as an image  |
' crawler to be integrated with google's image search.        |
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'
'______________________________________________________________
'                      Contact Information                     |
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|
' If you have any questions/comments/suggestions you           |
' may contact me in the following manners:                     |
' E-mail Address: theproductofisolation@hotmail.com            |
' Aol Instant Messenger: DarklohrIsolate                       |
' MSN Instant Messenger: theproductofisolation@hotmail.com     |
' Yahoo! Instant Messenger: lohrbasmetalaly                    |
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Status As String
Public Link As String

Sub SetStatus(LinksChecked As Long, LinksTotal As Long)
    ' Inputs:
    '         LinksChecked = How many links have been found on document
    '         LinksTotal   - How many links are in the document in total
    
    ' Result/Purpose:
    '                 Updates Form value (Status) with the current (Status)
    '                 on searching a webpage for links.
    
    Status = LinksChecked & Chr(1) & LinksTotal ' Sets Status (Links Done Checking/Total Links)
    DoEvents
End Sub

Function GetURLSource(URL As String)
    ' Input:
    '          URL = Link we want the source-code to
    ' Output:
    '          The source code of (URL)
    
    ' Purpose:
    '          This function returns the source code of the (URL) provided
    GetURLSource = frmMain.Inet1.OpenURL(URL)
End Function

Function SearchTag(Tag As String, LLink As String, Optional EndTag As String = ">") As Variant
    ' Inputs:
    '         Tag = The Tag we want to search for
    '         Link = The Link we want to search for our (Tag)
    '         EndTag = Optional feature which lets you get the body
    '                  of a tag.
    '                  Example:
    '                           <a href="blah.html">Blah</a>
    '                           Returns: Blah (needs body)
    '                  Since the default EndTag is set to ">"
    '                  then it will only go as far as
    '                  <a href="blah.html">
    '                  and you will not be able to get the
    '                  "Blah" because it is in the body of the tag.
    ' Ouput:
    '        Returns complete (Tag) occurances and properties on
    '        the link (Link)
    
    ' Purpose:
    '        This function returns (Tag) occurances in an array along
    '        with it's properties found on (Link). It has an optional
    '        (EndTag) attribute which lets you chose what the EndTag is
    '        because some HTML Properties are different in tagging.
    
    '        PLEASE DO NOT CHANGE (EndTag) ENLESS YOU KNOW WHAT YOU ARE DOING
    
    '        Examples:
    '                - EndTag set to ">":
    '                  It is good to set (EndTag) to ">" when you are looking for "<img" as a Tag
    '                  because it doesn't require you to end it in an "</img>" way so you probably
    '                  won't find it and the program'll go crazy
    '                  <img src="blah.jpg">
    '
    '                - EndTag set to "</a>"
    '                  It is good to set (EndTag) to "</a>" when you are looking for
    '                  the body of an "<a" tag such as looking for "Blah" in
    '                  "<a href="link.html">Blah</a>" which you can't find if
    '                  the program stops at ">"
    
    Dim TotalTags As Long, TagsChecked As Long, TagsRemaining As Long
    Dim Tags() As String, tmpTags As String
    Dim Source As String
    
    Source = GetURLSource(LLink)
    Link = LLink
    
    TotalTags = TotalInStr(1, Source, Tag, False) ' Count total (Tags)
    TagsChecked = 0 ' Tags Checked
    TagsRemaining = TotalTags - TagsChecked ' Tags Remaining
    
    SetStatus TagsChecked, TotalTags ' Set Status (0/Total Tags Completed)
    DoEvents
    If InStr(1, Source, Tag, vbTextCompare) <= 0 Then Exit Function ' No Tags Found
    
    Dim tmpSrc As String, i As Long, OpenTagPos As Long, CloseTagDifLen As Long
    tmpSrc = Mid(Source, InStr(1, Source, Tag, vbTextCompare)) ' Move Source up to first (Tag)
    
    tmpTags = LLink & Chr(1)
    
    Do While TagsChecked < TotalTags
        If TagsChecked >= TotalTags Then
            SetStatus TagsChecked, TotalTags ' Set Status (x/Total Tags Completed)
            DoEvents
            Exit Do ' No more (Tags) left to check
        End If
        
        OpenTagPos = InStr(1, tmpSrc, Tag, vbTextCompare) ' finds tag occurance
        CloseTagDifLen = InStr(1, tmpSrc, EndTag, vbTextCompare) - OpenTagPos + Len(EndTag) ' finds ">" (ending of tag occurance)
        
        tmpTags = tmpTags & Mid(tmpSrc, OpenTagPos, CloseTagDifLen) & Chr(1) ' Finds all (Tag) parameters
        
        TotalTags = TotalInStr(1, Source, Tag, False) ' Count total (Tags)
        TagsChecked = TotalInStr(1, tmpTags, Tag, False) ' Tags Checked
        TagsRemaining = TotalTags - TagsChecked ' Tags Remaining

        SetStatus TagsChecked, TotalTags ' Set Status (x/Total Tags Completed)
        DoEvents
        If TagsChecked >= TotalTags Then
            SetStatus TagsChecked, TotalTags ' Set Status (x/Total Tags Completed)
            DoEvents
            Exit Do ' No more (Tags) left to check
        End If
        
        tmpSrc = Mid(tmpSrc, InStr(OpenTagPos + CloseTagDifLen, tmpSrc, Tag)) ' moves to next (Tag) position
    Loop
    
    tmpTags = Mid(tmpTags, 1, Len(tmpTags) - 1)
    If Trim(tmpTags) = "" Then
        tmpTags = "Search returned no results." & Chr(1)
    End If
    Tags() = Split(tmpTags, Chr(1))
    SearchTag = Tags()
End Function

Function ViewProperty(BaseLink As String, Tag As String, Property As String) As String
    ' Inputs:
    '         Tag = Tag we want to search for (Property)
    '               Format:
    '                       <a href="musick.html">My Musick</a>
    '                       <a href="stuff.html">
    '                       <img src="hotgirl.jpg">
    '         Property = Property we want to extract from (Tag)
    '               Format:
    '                       <a href="musick.html">My Musick</a>
    '                            Set to "href" to return "musick.html"
    ' Output:
    '         Returns (Property) in (Tag)
    
    ' Purpose:
    '         This function returns (Property) in (Tag) with the special
    '         key called "BodyTag" where if set as the string of (Property)
    '         then the body of the string will be the output such as:
    '         <a href="musick.html">My Musick</a> -> Returns: "My Musick"
    Dim tmpTag As String, tmpProperty() As String
    
    If Property = "All" Then
        ViewProperty = Tag ' return complete tag
        Exit Function
    End If
    
    If Property = "BodyTag" Then ' They want the middle of the tag:
        ' Example: <a href="blah.html">BLAH</a>"
        ' Output: BLAH
        
        tmpTag = Mid(tmpTag, InStr(1, tmpTag, ">") + 1, InStr(1, tmpTag, "<", vbTextCompare) - InStr(1, tmpTag, ">") - 1)
        ViewProperty = tmpTag
        Exit Function
    End If
    
    tmpTag = Mid(Tag, InStr(1, Tag, " ") + 1) ' Remove opening tag "<img " or "<a " ect
    
    ' They don't want the text inbetween the tag, so
    ' we can delete everything after ">" of the first tag
    
    tmpTag = Mid(tmpTag, 1, InStr(1, tmpTag, ">") - 1) ' Finds end of first tag and deletes everything after it
    
    ' Now we have all the properties in such format:
    ' href=http://www.yahoo.com/directory method='post' onmouseover=""
    ' What we need to do is split those properties by " "
    
    ' Then the output will be:
    ' Array(0) = href=http://www.yahoo.com/directory; (1) = method='post' (2) = onmouseover=""
    
    ' Check the first (Len(Property)) of each string to match Property (no case matching)
    ' If it is equal, then we have found the right property so we should delete the tag itself
    ' and then check for single quotations or double quotations at the beginning and end
    ' of the string left
    
    ' Step 1) Split by " "
    tmpProperty() = Split(tmpTag, " ")
    
    ' Step 2) Compare property to wanted property
    Dim i As Long
    For i = LBound(tmpProperty()) To UBound(tmpProperty())
        If LCase(Mid(tmpProperty(i), 1, Len(Property))) = LCase(Property) Then
            ' LCase everything temporarily to have no case matching
            ' The two sides match so we have found the property we wanted
            
            ' First delete the property tag
            tmpProperty(i) = Mid(tmpProperty(i), Len(Property) + 2)
            
            ' Check for quotations
            If Mid(tmpProperty(i), 1, 1) = "'" Or Mid(tmpProperty(i), 1, 1) = Chr(34) Then ' Check for ' and " quotations at beginning of string
                ' Delete beginning quotation
                tmpProperty(i) = Mid(tmpProperty(i), 2)
            End If
            
            'Check for ending quotations
            If Right(tmpProperty(i), 1) = "'" Or Right(tmpProperty(i), 1) = Chr(34) Then ' Check for ' and " quotations at the end of the string
                ' Delete ending quotation
                tmpProperty(i) = Mid(tmpProperty(i), 1, Len(tmpProperty(i)) - 1)
            End If
            
            tmpTag = tmpProperty(i) ' Save needed string from Array
            Exit For
        End If
    Next i
    
    ' Check wheather the link is relative or direct
    Dim tmpBase As String ' String to hold the base incase link is relative
    
    If LCase(Left(tmpTag, Len("http:"))) = LCase("http:") Or LCase(Left(tmpTag, Len("ftp:"))) = LCase("ftp:") Or LCase(Left(tmpTag, Len("mailto:"))) = LCase("mailto:") Then
        ' Link is direct... leave it as it is and output
        ViewProperty = tmpTag
    Else
        ' Link is relative... add the base to it
        If Mid(tmpTag, 1, 1) = "/" Then
            tmpTag = Mid(tmpTag, 2)
        End If
        ViewProperty = BaseLink & tmpTag
    End If
End Function

Function GetLinkBase(Link As String) As String
    ' Input: "http://www.yahoo.com/News/firstpage.html"
    ' Output: "http://www.yahoo.com/News/"
    
    ' Purpose: This function returns the base of a link.
    '          Thus links pointing to files are changed
    '          to not contain the file but only the directory
    '          it is in.
    
    
    
    Dim tmpLink As String
    tmpLink = Mid(Link, InStrRev(Link, "/")) ' Finds last / and deletes all text before it
    
    ' Format might be:
    ' /link.html
    ' Should be:
    ' /
    
    If Len(tmpLink) > 1 And InStr(1, tmpLink, ".") > 0 Then
        ' Base is pointing to a file and not a directory
        ' Delete "/" and delete the last (Len(tmpLink)) of Link
        
        ' Delete "/"
        tmpLink = Right(tmpLink, Len(tmpLink) - 1)
        
        ' Delete last (Len(tmpLink)) of (Link)
        tmpLink = Mid(Link, 1, Len(Link) - Len(tmpLink))
    End If
    
    If LCase(tmpLink) = LCase("http://") Or LCase(tmpLink) = LCase("ftp://") Then
        ' The last "/" found was in "http://" or "ftp://"
        ' Thus you can figure out that it has no "/" at the end
        ' However it is a direct link, so all we need to do is add the "/"
        tmpLink = Link & "/"
    End If
    
    If tmpLink = "/" Then
        ' Only a "/" was found because the link is already pointing to a directory
        ' Thus we need not change anything
        tmpLink = Link
    End If
    
    If Len(tmpLink) > 1 And Left(tmpLink, 1) = "/" Then
        ' There is a "/" and a directory name following it
        ' Don't change anything but add the ending "/"
        tmpLink = Link & "/"
    End If
    
    GetLinkBase = tmpLink
End Function

Function TotalInStr(Start As Long, String1 As String, String2 As String, CaseMatch As Boolean) As Long
    ' Inputs:
    '          Start = (STARTS WITH 1 NOT 0) Tells the function where to start on the string to search for (String2)
    '          String1 = String to look for (String2) in
    '          String2 = String that is looked for in (String1)
    '          CaseMatch = (True/False) True = Case-sensitive; False = Case-insensitive;
    
    ' Output:
    '          Input: Test = TotalInStr(1, "I wAnt to know how mAny A's Are in this piece of text", "a", False)
    '          Output: 4 (4 a's were found in the case-insensitive search)
    
    ' Purpose:
    '          This function finds the total number of occurances of (String2) in (String1)
    '          with options of where to start on (String1) by setting (Start) and wheather
    '          to be case-sensitive or insensitive.
    
    If Start <= 0 Then
        Start = 1 ' Error prevention incase user enters a number equal or less than 0 to 1
    End If
    
    Dim tmpString1 As String, tmpString2 As String
    tmpString1 = Mid(tmpString1, Start) ' Start on (Start)
    
    If CaseMatch = False Then ' Converts temp. strings to lowercase (no case match)
        tmpString1 = LCase(String1)
        tmpString2 = LCase(String2)
    End If
    
    TotalInStr = UBound(Split(String1, String2)) ' Split (String1) by (String2) to count all
                                                 ' occurances between (String2) in (String1)
End Function
