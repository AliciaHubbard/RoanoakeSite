<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' *** Logout the current user.
MM_logoutRedirectPage = "loggedout.asp"
Session.Contents.Remove("MM_Username")
Session.Contents.Remove("MM_UserAuthorization")
If (MM_logoutRedirectPage <> "") Then Response.Redirect(MM_logoutRedirectPage)
%>
<!--#include file="Connections/roanoake.asp" -->
<%
Dim rs_logout
Dim rs_logout_cmd
Dim rs_logout_numRows

Set rs_logout_cmd = Server.CreateObject ("ADODB.Command")
rs_logout_cmd.ActiveConnection = MM_roanoake_STRING
rs_logout_cmd.CommandText = "SELECT * FROM tbl_login" 
rs_logout_cmd.Prepared = true

Set rs_logout = rs_logout_cmd.Execute
rs_logout_numRows = 0
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
<p>Logout</p>
<p>Logging you out</p>
<!-- InstanceEndEditable -->
</div>
</body>
<!-- InstanceEnd --></html>
<%
rs_logout.Close()
Set rs_logout = Nothing
%>
