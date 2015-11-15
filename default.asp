<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/roanoake.asp" -->
<%
Dim alphabetical__MMColParam
alphabetical__MMColParam = "1"
If (Request("MM_EmptyValue") <> "") Then 
  alphabetical__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim alphabetical
Dim alphabetical_cmd
Dim alphabetical_numRows

Set alphabetical_cmd = Server.CreateObject ("ADODB.Command")
alphabetical_cmd.ActiveConnection = MM_roanoake_STRING
alphabetical_cmd.CommandText = "SELECT a_name, a_description, a_admission, type FROM tbl_attractions WHERE approved = ? ORDER BY a_name ASC" 
alphabetical_cmd.Prepared = true
alphabetical_cmd.Parameters.Append alphabetical_cmd.CreateParameter("param1", 200, 1, 255, alphabetical__MMColParam) ' adVarChar

Set alphabetical = alphabetical_cmd.Execute
alphabetical_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 5
Repeat1__index = 0
alphabetical_numRows = alphabetical_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim alphabetical_total
Dim alphabetical_first
Dim alphabetical_last

' set the record count
alphabetical_total = alphabetical.RecordCount

' set the number of rows displayed on this page
If (alphabetical_numRows < 0) Then
  alphabetical_numRows = alphabetical_total
Elseif (alphabetical_numRows = 0) Then
  alphabetical_numRows = 1
End If

' set the first and last displayed record
alphabetical_first = 1
alphabetical_last  = alphabetical_first + alphabetical_numRows - 1

' if we have the correct record count, check the other stats
If (alphabetical_total <> -1) Then
  If (alphabetical_first > alphabetical_total) Then
    alphabetical_first = alphabetical_total
  End If
  If (alphabetical_last > alphabetical_total) Then
    alphabetical_last = alphabetical_total
  End If
  If (alphabetical_numRows > alphabetical_total) Then
    alphabetical_numRows = alphabetical_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (alphabetical_total = -1) Then

  ' count the total records by iterating through the recordset
  alphabetical_total=0
  While (Not alphabetical.EOF)
    alphabetical_total = alphabetical_total + 1
    alphabetical.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (alphabetical.CursorType > 0) Then
    alphabetical.MoveFirst
  Else
    alphabetical.Requery
  End If

  ' set the number of rows displayed on this page
  If (alphabetical_numRows < 0 Or alphabetical_numRows > alphabetical_total) Then
    alphabetical_numRows = alphabetical_total
  End If

  ' set the first and last displayed record
  alphabetical_first = 1
  alphabetical_last = alphabetical_first + alphabetical_numRows - 1
  
  If (alphabetical_first > alphabetical_total) Then
    alphabetical_first = alphabetical_total
  End If
  If (alphabetical_last > alphabetical_total) Then
    alphabetical_last = alphabetical_total
  End If

End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = alphabetical
MM_rsCount   = alphabetical_total
MM_size      = alphabetical_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
alphabetical_first = MM_offset + 1
alphabetical_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (alphabetical_first > MM_rsCount) Then
    alphabetical_first = MM_rsCount
  End If
  If (alphabetical_last > MM_rsCount) Then
    alphabetical_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE HTML>
<html><!-- InstanceBegin template="/Templates/roanoake_events.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<meta charset="utf-8">
<meta name="viewport" content="initial-scale=1.0, width=device-width" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Roanoke</title>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
<script src="js/modernizr.js"></script>
<!--[if lt IE 9]>
<script src="http://html5shim.googlecode.com/svn/trunk/html5.js"></script>
<![endif]-->
<link href="css/rstyles.css" rel="stylesheet" type="text/css">
</head>
<body class="container">
<header>
<h1>Roanoke Attractions</h1>
</header>
<nav>
  <li>
    <p><a href="default.asp" title="Alphabetical List">Alphabetical List<img src="images/star_small.jpg"></a></p>
  </li>
  <li><p><a href="search.asp" title="Search by Type">Search by Type<img src="images/star_small.jpg"></a></p></li>
  <li><p><a href="add.asp" title="Add an Event">Add an Attraction<img src="images/star_small.jpg"></a></p></li>
  <li><p><a href="star.asp" title="Story of the Star">Story of the Star<img src="images/star_small.jpg"></a></p></li>
  <li><p><a href="approve.asp" title="Administrator">Administrator<img src="images/star_small.jpg"></a></p></li>
</nav>
<div id=container>
<!-- InstanceBeginEditable name="page heading" -->
  <h2>Alphabetical List</h2>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="content" -->
<table>
  <tr>
    <td>Name</td>
    <td>Description</td>
    <td>Admission</td>
    <td>Type</td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT alphabetical.EOF)) %>
    <tr>
      <td><%=(alphabetical.Fields.Item("a_name").Value)%></td>
      <td><%=(alphabetical.Fields.Item("a_description").Value)%></td>
      <td><%=(alphabetical.Fields.Item("a_admission").Value)%></td>
      <td><%=(alphabetical.Fields.Item("type").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  alphabetical.MoveNext()
Wend
%>
</table>
<A HREF="<%=MM_moveFirst%>" class="leftalign">First</A>&nbsp; &nbsp;
<A HREF="<%=MM_movePrev%>" class="leftalign">Previous</A> 
<A HREF="<%=MM_moveLast%>" class="rightalign">Last</A>&nbsp; &nbsp;
<A HREF="<%=MM_moveNext%>"class="rightalign">Next</A>
<p class="center">Showing <%=(alphabetical_first)%> to <%=(alphabetical_last)%> of <%=(alphabetical_total)%> </p>
 <!-- InstanceEndEditable -->
</div>
</body>
<!-- InstanceEnd --></html>
<%
alphabetical.Close()
Set alphabetical = Nothing
%>
