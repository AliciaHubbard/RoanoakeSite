<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/roanoake.asp" -->
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
    MM_editRedirectUrl = "thankyou.asp"
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
Dim rs_attractions
Dim rs_attractions_cmd
Dim rs_attractions_numRows

Set rs_attractions_cmd = Server.CreateObject ("ADODB.Command")
rs_attractions_cmd.ActiveConnection = MM_roanoake_STRING
rs_attractions_cmd.CommandText = "SELECT * FROM tbl_attractions" 
rs_attractions_cmd.Prepared = true

Set rs_attractions = rs_attractions_cmd.Execute
rs_attractions_numRows = 0
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
  <h2>Add An Attraction</h2>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="content" -->
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
      <td><label for="type"></label>
        <select name="type" id="type">
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
        </select>
        <input name="approved" type="hidden" id="approved" value="0"></td>
    </tr>
  </table>
  <p>
    <input type="submit" value="Insert record">
    <input type="hidden" name="MM_insert" value="form1">
  </p>
</form>
<p>Events will be reviewed before they are visible (sorry  wiseguys)</p>


<!-- InstanceEndEditable -->
</div>
</body>
<!-- InstanceEnd --></html>
<%
rs_attractions.Close()
Set rs_attractions = Nothing
%>
<%
rs_type.Close()
Set rs_type = Nothing
%>
