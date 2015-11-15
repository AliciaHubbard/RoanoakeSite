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
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_roanoake_STRING
    MM_editCmd.CommandText = "INSERT INTO tbl_attractions (a_name, a_description, a_admission, type, approved) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 255, Request.Form("a_name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 255, Request.Form("a_description")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 255, Request.Form("a_admission")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 255, Request.Form("type")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("approved"), Request.Form("approved"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "added.asp"
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
Dim rs_add
Dim rs_add_cmd
Dim rs_add_numRows

Set rs_add_cmd = Server.CreateObject ("ADODB.Command")
rs_add_cmd.ActiveConnection = MM_roanoake_STRING
rs_add_cmd.CommandText = "SELECT * FROM tbl_attractions" 
rs_add_cmd.Prepared = true

Set rs_add = rs_add_cmd.Execute
rs_add_numRows = 0
%>
<%
Dim rs_type
Dim rs_type_cmd
Dim rs_type_numRows

Set rs_type_cmd = Server.CreateObject ("ADODB.Command")
rs_type_cmd.ActiveConnection = MM_roanoake_STRING
rs_type_cmd.CommandText = "SELECT * FROM tbl_type" 
rs_type_cmd.Prepared = true

Set rs_type = rs_type_cmd.Execute
rs_type_numRows = 0
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
<p>Add an Attraction
</p>
<form method="POST" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap align="right">Name:</td>
      <td><input type="text" name="a_name" value="" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Description:</td>
      <td><input type="text" name="a_description" value="" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Admission:</td>
      <td><input type="text" name="a_admission" value="" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Type:</td>
      <td><select name="type" id="type">
        <%
While (NOT rs_type.EOF)
%>
        <option value="<%=(rs_type.Fields.Item("type").Value)%>"><%=(rs_type.Fields.Item("type").Value)%></option>
        <%
  rs_type.MoveNext()
Wend
If (rs_type.CursorType > 0) Then
  rs_type.MoveFirst
Else
  rs_type.Requery
End If
%>
      </select></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">&nbsp;</td>
      <td><input type="submit" value="Insert record"></td>
    </tr>
  </table>
  <input type="hidden" name="approved" value="1">
  <input type="hidden" name="MM_insert" value="form1">
</form>
<p>&nbsp;</p>
<!-- InstanceEndEditable -->
</div>
</body>
<!-- InstanceEnd --></html>
<%
rs_add.Close()
Set rs_add = Nothing
%>
<%
rs_type.Close()
Set rs_type = Nothing
%>
