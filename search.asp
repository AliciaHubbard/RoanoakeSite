<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/roanoake.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_roanoake_STRING
Recordset1_cmd.CommandText = "SELECT * FROM tbl_type" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
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
  <h2>Attractions by Type</h2>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="content" -->
<form name="form1" method="post" action="type.asp">
  <label for="type"></label>
  <select name="type" id="type">
    <%
While (NOT Recordset1.EOF)
%>
    <option value="<%=(Recordset1.Fields.Item("type").Value)%>"><%=(Recordset1.Fields.Item("type").Value)%></option>
    <%
  Recordset1.MoveNext()
Wend
If (Recordset1.CursorType > 0) Then
  Recordset1.MoveFirst
Else
  Recordset1.Requery
End If
%>
  </select>
  <input type="submit" name="button" id="button" value="Search">
</form>
<!-- InstanceEndEditable -->
</div>
</body>
<!-- InstanceEnd --></html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
