<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/roanoake.asp" -->
<%
Dim rs_login
Dim rs_login_cmd
Dim rs_login_numRows

Set rs_login_cmd = Server.CreateObject ("ADODB.Command")
rs_login_cmd.ActiveConnection = MM_roanoake_STRING
rs_login_cmd.CommandText = "SELECT * FROM tbl_login" 
rs_login_cmd.Prepared = true

Set rs_login = rs_login_cmd.Execute
rs_login_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "approve.asp"
  MM_redirectLoginFailed = "access_denied.asp"

  MM_loginSQL = "SELECT username, passowrd"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM tbl_login WHERE username = ? AND passowrd = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_roanoake_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 255, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 255, Request.Form("password")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
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
  <h2>Administrator Login</h2>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="content" -->
<form name="form1" method="POST" action="<%=MM_LoginAction%>">
  <p>Username:
  <input type="text" name="username" id="username">
    &nbsp; &nbsp;Password: 
    <input type="text" name="password" id="password">
  </p>
  <p>
    <input type="submit" name="button" id="button" value="Login">
  </p>
</form>
<!-- InstanceEndEditable -->
</div>
</body>
<!-- InstanceEnd --></html>
<%
rs_login.Close()
Set rs_login = Nothing
%>
