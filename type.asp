<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/roanoake.asp" -->
<%
Dim rs_roanoke__MMColParam
rs_roanoke__MMColParam = "1"
If (Request.Form("type") <> "") Then 
  rs_roanoke__MMColParam = Request.Form("type")
End If
%>
<%
Dim rs_roanoke
Dim rs_roanoke_cmd
Dim rs_roanoke_numRows

Set rs_roanoke_cmd = Server.CreateObject ("ADODB.Command")
rs_roanoke_cmd.ActiveConnection = MM_roanoake_STRING
rs_roanoke_cmd.CommandText = "SELECT a_name, a_description, a_admission, type FROM tbl_attractions WHERE (type = ?) AND (approved = 1)" 
rs_roanoke_cmd.Prepared = true
rs_roanoke_cmd.Parameters.Append rs_roanoke_cmd.CreateParameter("param1", 200, 1, 255, rs_roanoke__MMColParam) ' adVarChar

Set rs_roanoke = rs_roanoke_cmd.Execute
rs_roanoke_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rs_roanoke_numRows = rs_roanoke_numRows + Repeat1__numRows
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
<table>
  <tr>
    <td>Name</td>
    <td>Description</td>
    <td>Admission</td>
    <td>Type</td>
  </tr>
  <% While ((Repeat1__numRows <> 0) AND (NOT rs_roanoke.EOF)) %>
    <tr>
      <td><%=(rs_roanoke.Fields.Item("a_name").Value)%></td>
      <td><%=(rs_roanoke.Fields.Item("a_description").Value)%></td>
      <td><%=(rs_roanoke.Fields.Item("a_admission").Value)%></td>
      <td><%=(rs_roanoke.Fields.Item("type").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs_roanoke.MoveNext()
Wend
%>
</table>
<!-- InstanceEndEditable -->
</div>
</body>
<!-- InstanceEnd --></html>
<%
rs_roanoke.Close()
Set rs_roanoke = Nothing
%>
