<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/roanoake.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
Dim rs_approve__MMColParam
rs_approve__MMColParam = "0"
If (Request("MM_EmptyValue") <> "") Then 
  rs_approve__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rs_approve
Dim rs_approve_cmd
Dim rs_approve_numRows

Set rs_approve_cmd = Server.CreateObject ("ADODB.Command")
rs_approve_cmd.ActiveConnection = MM_roanoake_STRING
rs_approve_cmd.CommandText = "SELECT * FROM tbl_attractions WHERE approved = ?" 
rs_approve_cmd.Prepared = true
rs_approve_cmd.Parameters.Append rs_approve_cmd.CreateParameter("param1", 5, 1, -1, rs_approve__MMColParam) ' adDouble

Set rs_approve = rs_approve_cmd.Execute
rs_approve_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rs_approve_numRows = rs_approve_numRows + Repeat1__numRows
%>
<%
Dim MM_paramName 
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
<!DOCTYPE HTML>
<html><!-- InstanceBegin template="/Templates/Administrator.dwt.asp" codeOutsideHTMLIsLocked="false" -->
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
  <h2>Administrator Dashboard</h2>
  <a href="approve.asp" title="Approve">Approve</a> &nbsp;  &nbsp;
  <a href="edit.asp" title="Edit/Delete">Edit/Delete</a>&nbsp;  &nbsp;  
  <a href="admin_add.asp" title="Add">Add</a> &nbsp;  &nbsp;
  <a href="logout.asp" title="Logout">Logout</a>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="content" -->
<p>Review Submissions
</p>
<table>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>Name</td>
    <td>Description</td>
    <td>Admission</td>
    <td>Type</td>
    <td>Approved</td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rs_approve.EOF)) %>
    <tr>
      <td><A HREF="reviewdelete.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "a_name=" & rs_approve.Fields.Item("a_name").Value %>">Delete</A></td>
      <td><A HREF="reviewapprove.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "a_name=" & rs_approve.Fields.Item("a_name").Value %>">Approve</A></td>
      <td><%=(rs_approve.Fields.Item("a_name").Value)%></td>
      <td><%=(rs_approve.Fields.Item("a_description").Value)%></td>
      <td><%=(rs_approve.Fields.Item("a_admission").Value)%></td>
      <td><%=(rs_approve.Fields.Item("type").Value)%></td>
      <td><%=(rs_approve.Fields.Item("approved").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_approve.MoveNext()
Wend
%>
</table>
<!-- InstanceEndEditable -->
</div>
</body>
<!-- InstanceEnd --></html>
<%
rs_approve.Close()
Set rs_approve = Nothing
%>
