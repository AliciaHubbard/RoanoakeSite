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
Dim rs_delete__MMColParam
rs_delete__MMColParam = "1"
If (Request.QueryString("a_name") <> "") Then 
  rs_delete__MMColParam = Request.QueryString("a_name")
End If
%>
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_roanoake_STRING
    MM_editCmd.CommandText = "DELETE FROM tbl_attractions WHERE a_name = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 200, 1, 255, Request.Form("MM_recordId")) ' adVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

     'append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "edit.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If

End If
%>
<%
Dim rs_delete
Dim rs_delete_cmd
Dim rs_delete_numRows

Set rs_delete_cmd = Server.CreateObject ("ADODB.Command")
rs_delete_cmd.ActiveConnection = MM_roanoake_STRING
rs_delete_cmd.CommandText = "SELECT * FROM tbl_attractions WHERE a_name = ?" 
rs_delete_cmd.Prepared = true
rs_delete_cmd.Parameters.Append rs_delete_cmd.CreateParameter("param1", 200, 1, 255, rs_delete__MMColParam) ' adVarChar

Set rs_delete = rs_delete_cmd.Execute
rs_delete_numRows = 0
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
<p>Delete A Record</p>
<form name="form1" method="POST" action="<%=MM_editAction%>">
<table border="0" cellspacing="0" cellpadding="10">
  <tr>
    <th>Name</th>
    <th>Description</th>
  </tr>
  <tr>
    <td><%=(rs_delete.Fields.Item("a_name").Value)%></td>
    <td><%=(rs_delete.Fields.Item("a_description").Value)%></td>
  </tr>
</table>
<p>
  <input type="submit" name="button" id="button" value="Delete">
</p>
<input type="hidden" name="MM_delete" value="form1">
<input type="hidden" name="MM_recordId" value="<%= rs_delete.Fields.Item("a_name").Value %>">
</form>
<p>&nbsp;</p>
<p></p>
<!-- InstanceEndEditable -->
</div>
</body>
<!-- InstanceEnd --></html>
<%
rs_delete.Close()
Set rs_delete = Nothing
%>
