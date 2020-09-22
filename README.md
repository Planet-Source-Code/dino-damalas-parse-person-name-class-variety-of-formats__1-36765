<div align="center">

## Parse Person Name Class \(variety of formats\)


</div>

### Description

Got extremely tired trying to find a quick and systematic way of parsing a field that contained a user's name in a variety of formats, so I created this little class that will parse out a person's name into first, middle, last, title, prefix, suffix. It can handle names like Dr. John Doe - Dr. Doe, John P - Doe, John - John P. Doe, Jr. - and a few more formats. Hope others will find this useful. Currently the class cannot handle muliple suffixes and multiple titles. If someone reworks it to make it better, please send it my way. FYI- commented all over, should be easy to read.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dino Damalas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dino-damalas.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dino-damalas-parse-person-name-class-variety-of-formats__1-36765/archive/master.zip)





### Source Code

```
'****************************************************************************
' Module Name:   clsNameParse
' Module Type:   Class Module
' Filename:     clsNameParse.cls
' Author:      Dino Damalas
' Date:       7/10/2002
' References:    Microsoft Regular Expression Object
' Purpose:     Use this class when dealing with inconsistent
'          person name formats. This object will parse
'          a person's name into
'            - Prefix
'            - Suffix
'            - First Name
'            - Middle Name / Middle Initial
'            - Last Name
'            - Title
'          examples: Dr. John P Doe Jr
'               Dr. Doe, John P.
'               John Doe
'               Doe, John P.
'               John P. Doe, CEO
'               ...etc
'
' Example Use:   Dim objParse as new clsParse
'          objParse.ParseName("Dr. Doe, John P.")
'          strFirstName  = objParse.FirstName
'          strLastName   = objParse.LastName
'          strMiddleName  = objParse.MiddleName
'          strMiddleInit  = objParse.MiddleInitial
'          strPrefix    = objParse.Prefix
'          strSuffix    = objParse.Suffix
'          strTitle    = objparse.title
'          set objParse = nothing
'
'*****************************************************************************
Option Explicit
'--member var declaration
Private mobjRegExp As RegExp
Private mstrPrefix As String
Private mstrSuffix As String
Private mstrLastName As String
Private mstrFirstName As String
Private mstrMiddleName As String
Private mstrMiddleInitial As String
Private mstrTitle As String
Private mstrFullName As String
Private mblnHasError As Boolean
Private mstrErrorMessage As String
'===============================================================================
' Name:   Class_Initialize
' Input:  None
' Output:  None
' Purpose: Initialize Class - initialize a few vars and objects
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Private Sub Class_Initialize()
  mblnHasError = False
  mstrErrorMessage = ""
  Set mobjRegExp = New RegExp
  mobjRegExp.IgnoreCase = True
End Sub
'===============================================================================
' Name:   Class_Terminate
' Purpose: Clean up.. destory regexp object
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Private Sub Class_Terminate()
  Set mobjRegExp = Nothing
End Sub
'===============================================================================
' Name:   ParseName
' Input:
'      strName - String :: A persons full name
' Output:
'      none
' Purpose: Main sub to initiate parsing of name
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Public Sub ParseName(ByVal strName As String)
  '-- pick apart name by removing prefix, suffix, and title
  strName = Trim(fncExtractSuffix(strName))
  strName = Trim(fncExtractPrefix(strName))
  strName = Trim(fncExtractTitle(strName))
  mobjRegExp.Global = True
  '-- check for last, first combo (Doe, John) ----
  mobjRegExp.Pattern = "[^ \f\n\r\t\v\,]+\,\s+\S+ "
  If mobjRegExp.Test(strName) = True Then
    Call subParseLastFirst(strName)
  Else
    '-- check if first middle last combo (John P. Doe) ---
    mobjRegExp.Pattern = "^\S+\s+\S+\s+\S+$"
    If mobjRegExp.Test(strName) Then
      Call subParseFirstMiddleLast(strName)
    Else
      '-- check if first last combo (John Doe) --
      mobjRegExp.Pattern = "^\S+\s+\S+$"
      If mobjRegExp.Test(strName) Then
        Call subParseFirstLast(strName)
      Else
        '--if does not fit in this format tell user we have a prob
        mblnHasError = True
        mstrErrorMessage = "Unable To Parse"
      End If
    End If
  End If
End Sub
'===============================================================================
' Name:   fncExtractPrefix
' Input:
'      strName - String :: Person's Full Name
' Output:
'      String :: Name without prefix
' Purpose: Removes the prefix from the name and sets the Prefix property of the class
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Private Function fncExtractPrefix(ByVal strName As String) As String
  '--declare vars
  Dim aryPrefix As Variant
  Dim intCounter As Integer
  Dim strReturn As String
  Dim objMatches As MatchCollection
  '--initialize vars
  strReturn = strName
  '--populate array with a bunch of possible prefixes
  aryPrefix = Array("mr", "mrs", "miss", "dr", "prof", "pvt", "pfc", "lcpl", "cpl", "spc", "sgt", "ssgt", "gysgt", "msgt", "mgysgt", "lt", "capt", "col", "ltcol", "gen", "adm", "rdm")
  '--loop through the array looking for matches using regexp
  mobjRegExp.Global = False
  For intCounter = 0 To UBound(aryPrefix)
    mobjRegExp.Pattern = "^" & aryPrefix(intCounter) & "\.?\s+"
    If mobjRegExp.Test(strName) Then
      '-- if found, replace with empty string
      strReturn = Trim(mobjRegExp.Replace(strName, ""))
      Set objMatches = mobjRegExp.Execute(strName)
      '--set prefix property
      Me.Prefix = Trim(objMatches(0).Value)
      Set objMatches = Nothing
      Exit For
    End If
  Next
  fncExtractPrefix = strReturn
End Function
'===============================================================================
' Name:   fncExtractSuffix
' Input:
'      strName - String :: Person's Full Name
' Output:
'      String :: Name without suffix
' Purpose: Removes the suffix from the name and sets the Suffix property of the class
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Private Function fncExtractSuffix(ByVal strName As String) As String
  '--declare vars
  Dim arySuffix As Variant
  Dim intCounter As Integer
  Dim strReturn As String
  Dim objMatches As MatchCollection
  '--initialize vars
  strReturn = strName
  '--populate array with a bunch of possible suffixes
  arySuffix = Array("md", "i", "ii", "iid", "iii", "iv", "jr", "sr", "v", "vi", "vii", "do", "dds", "np", "pa", "phd", "ph d", "esq")
  '--loop through the array looking for matches using regexp
  mobjRegExp.Global = False
  For intCounter = 0 To UBound(arySuffix)
    mobjRegExp.Pattern = "\b" & arySuffix(intCounter) & "\.?(\s+|$)"
    If mobjRegExp.Test(strName) Then
       '-- if found, replace with empty string
      strReturn = Trim(mobjRegExp.Replace(strName, ""))
      Set objMatches = mobjRegExp.Execute(strName)
      '--set prefix property
      Me.Suffix = Trim(objMatches(0).Value)
      Set objMatches = Nothing
      Exit For
    End If
  Next
  fncExtractSuffix = strReturn
End Function
'===============================================================================
' Name:   fncExtractTitle
' Input:
'      strName - String :: Persons full name
' Output:
'      string :: Name without title
' Purpose: Removes title from name and sets the title property of the class
' Remarks: issues here.. if title is not behind a comma this will not work
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Private Function fncExtractTitle(ByVal strName As String) As String
  '--delcare vars
  Dim strReturn As String
  Dim intCommaPos As Integer
  Dim objMatches As MatchCollection
  Dim objMatch As Match
  '--initialize vars
  strReturn = strName
  '--get the first position of a comma
  intCommaPos = InStr(1, strName, ",", vbTextCompare)
  '--see if we have a comma in the name
  If intCommaPos > 0 Then
    mobjRegExp.Pattern = "[^ \f\n\r\t\v\,]+\,\s+\S+"
    '--check to see if this comma is lastname, firstname format
    If mobjRegExp.Test(strName) = True Then
      '--check to see if there is another comma since first is a last, first name seperator
      If InStr(intCommaPos + 1, strName, ",", vbTextCompare) > 0 Then
        '--if the last character is not a comma then parse out the title
        If Right(Trim(strName), 1) <> "," Then
          mobjRegExp.Pattern = "\,\s+\S+\s*$"
          Set objMatches = mobjRegExp.Execute(strName)
          For Each objMatch In objMatches
            '--set the title
            Me.Title = fncScrubString(objMatch.Value)
          Next
          Set objMatches = Nothing
          strReturn = mobjRegExp.Replace(strName, "")
        End If
      End If
    End If
  End If
  fncExtractTitle = strReturn
End Function
'===============================================================================
' Name:   fncScrubString
' Input:
'      strNamePart - String :: any name part - first last etc
' Output:
'      string - cleaned up version
' Purpose: removes any commas or extra spacings from name part
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Private Function fncScrubString(ByVal strNamePart As String) As String
  fncScrubString = Trim(Replace(strNamePart, ",", ""))
End Function
'===============================================================================
' Name:   subParseLastFirst
' Input:
'      strName - String :: Name without prefix, suffix, or title
' Purpose: Parses a name that is in LastName, FirstName format
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Private Sub subParseLastFirst(ByVal strName As String)
  '--declare vars
  Dim objMatches As MatchCollection
  Dim objMatch As Match
  Dim intCounter As Integer
  '--initialize
  intCounter = 1
  mobjRegExp.Global = True
  mobjRegExp.Pattern = "\S+"
  Set objMatches = mobjRegExp.Execute(strName)
  For Each objMatch In objMatches
    Select Case intCounter
      Case 1 '-- first time around is last name
        Me.LastName = fncScrubString(objMatch.Value)
      Case 2 '-- second time around is first name
        Me.FirstName = fncScrubString(objMatch.Value)
      Case 3 '-- if there is a third than its the middlename
        Me.MiddleName = fncScrubString(objMatch.Value)
    End Select
    intCounter = intCounter + 1
  Next
End Sub
'===============================================================================
' Name:   subParseFirstLast
' Input:
'      strName - String :: Name without prefix, suffix, or title
' Purpose: Parses a name in FirstName LastName format (no middle name)
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Private Sub subParseFirstLast(ByVal strName As String)
  '--declare vars
  Dim objMatches As MatchCollection
  Dim objMatch As Match
  Dim intCounter As Integer
  '--initialize
  intCounter = 1
  '--set up regexp object
  mobjRegExp.Global = True
  mobjRegExp.Pattern = "\S+"
  Set objMatches = mobjRegExp.Execute(strName)
  '--run through matches
  For Each objMatch In objMatches
    Select Case intCounter
      Case 1 '-- first time around we set first name
        Me.FirstName = fncScrubString(objMatch.Value)
      Case 2 '-- second time we set last name
        Me.LastName = fncScrubString(objMatch.Value)
    End Select
    intCounter = intCounter + 1
  Next
End Sub
'===============================================================================
' Name:   subParseFirstMiddleLast
' Input:
'      strName - String :: Name without prefix, suffix, or title
' Purpose: Parses a name in FirstName Middlename LastName format
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Private Sub subParseFirstMiddleLast(ByVal strName As String)
  '--declare vars
  Dim objMatches As MatchCollection
  Dim objMatch As Match
  Dim intCounter As Integer
  '--initialize vars
  intCounter = 1
  '--set up regexp object
  mobjRegExp.Global = True
  mobjRegExp.Pattern = "\S+"
  Set objMatches = mobjRegExp.Execute(strName)
  '--loop thorough matches
  For Each objMatch In objMatches
    Select Case intCounter
      Case 1 '-- first time is firstname
        Me.FirstName = fncScrubString(objMatch.Value)
      Case 2 '-- second time around is middlename
        Me.MiddleName = fncScrubString(objMatch.Value)
      Case 3 '-- last time around is last name
        Me.LastName = fncScrubString(objMatch.Value)
    End Select
    intCounter = intCounter + 1
  Next
End Sub
'===============================================================================
' Name:   Clear
' Purpose: Use this sub to clear out members when you implementing
'      in code where you don't reinstantiate the object again
' Author:  Dino Damalas
' Date:   7/10/2002
'===============================================================================
Public Sub Clear()
  Me.FirstName = ""
  Me.MiddleInitial = ""
  Me.MiddleName = ""
  Me.LastName = ""
  Me.Suffix = ""
  Me.Prefix = ""
  Me.Title = ""
  mblnHasError = False
  mstrErrorMessage = ""
End Sub
Public Property Get Prefix() As String
  Prefix = mstrPrefix
End Property
Public Property Let Prefix(ByVal strPrefix As String)
  mstrPrefix = strPrefix
End Property
Public Property Get Suffix() As String
  Suffix = mstrSuffix
End Property
Public Property Let Suffix(ByVal strSuffix As String)
  mstrSuffix = strSuffix
End Property
Public Property Get LastName() As String
  LastName = mstrLastName
End Property
Public Property Let LastName(ByVal strLastName As String)
  mstrLastName = strLastName
End Property
Public Property Get FirstName() As String
  FirstName = mstrFirstName
End Property
Public Property Let FirstName(ByVal strFirstName As String)
  mstrFirstName = strFirstName
End Property
Public Property Get MiddleName() As String
  MiddleName = mstrMiddleName
End Property
Public Property Let MiddleName(ByVal strMiddleName As String)
  mstrMiddleName = strMiddleName
  '--set up middle initial while we're here
  If Len(strMiddleName) > 1 Then
    Me.MiddleInitial = Left(strMiddleName, 1)
  Else
    Me.MiddleInitial = ""
  End If
End Property
Public Property Get MiddleInitial() As String
  MiddleInitial = mstrMiddleInitial
End Property
Public Property Let MiddleInitial(ByVal strMiddleInitial As String)
  mstrMiddleInitial = strMiddleInitial
End Property
Public Property Get Title() As String
  Title = mstrTitle
End Property
Public Property Let Title(ByVal strTitle As String)
  mstrTitle = strTitle
End Property
Public Property Get HasError() As Boolean
  HasError = mblnHasError
End Property
Public Property Get ErrorMessage() As String
  ErrorMessage = mstrErrorMessage
End Property
```

