<%if session("admin")=true then%>
<!--#include file="adovbs.inc"-->
<!--#include file="config.asp"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>工具栏</title>
<base target="main">
<link rel="stylesheet" href="master.css" type="text/css">
</head>

<body bgcolor="#DEF3FE">
<table width="100" border="1" cellpadding="0" cellspacing="0">
  <tr> 
    <td class="detaillabels" height="14"> 
      <div align="left"><a href="main.asp">主目录</a></div>
    </td>
  </tr>
  <tr> 
    <td class="detaillabels"> 
      <div align="left"><font color="#000000">表：</font></div>
    </td>
  </tr>
  <%
Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr
set objSchema = Con.OpenSchema(adSchemaTables)
Do While Not objSchema.EOF
	if objSchema("TABLE_TYPE") = "TABLE" then
%>
  <tr> 
    <td class="detaillabels"> 
      <div align="left">&nbsp;<a href="tableview.asp?table=<%=objSchema("TABLE_NAME")%>"><%=objSchema("TABLE_NAME")%></a></div>
    </td>
  </tr>
  <%
	end if
objSchema.MoveNext
Loop
%>
  <tr> 
    <td class="detaillabels"> 
      <div align="left"><font color="#000000">视图：</font></div>
    </td>
  </tr>
  <%
objSchema.MoveFirst
Do WHILE NOT objSchema.EOF
	if objSchema("TABLE_TYPE")="VIEW" then
%>
  <tr> 
    <td class="detaillabels"> 
		<a href="showview.asp?table=<%=objSchema("TABLE_NAME")%>&sql=SELECT+*+FROM+[<%=objSchema("TABLE_NAME")%>]"><%=objSchema("TABLE_NAME")%></a>
    </td>
  </tr>
  <%
	end if
objSchema.MoveNext
Loop
objSchema.Close
set objSchema=nothing
%>
  <tr> 
    <td class="detaillabels"> 
      <div align="left"><font color="#000000">存储过程：</font></div>
    </td>
  </tr>
  <%
set objSchema = Con.OpenSchema(adSchemaProcedures)
Do WHILE NOT objSchema.EOF
%>
  <tr> 
    <td class="detaillabels"> 
	<a href="showproc.asp?table=<%=objSchema("PROCEDURE_NAME")%>"><%=objSchema("PROCEDURE_NAME")%></a>
    </td>
  </tr>
  <%
objSchema.MoveNext
Loop
objSchema.Close
Set objSchema=nothing
Con.Close
Set Con=nothing
%>
  <tr> 
    <td class="detaillabels"> 
      <div align="left"><font color="#000000">其他功能：</font></div>
    </td>
  </tr>
  <tr> 
    <td class="detaillabels">&nbsp;<a href="compact.asp">压缩数据库</a></td>
  </tr>
  <tr> 
    <td class="detaillabels"> 
      <div align="left">&nbsp;<a href="chgpwd.asp">更改密码</a></div>
    </td>
  </tr>
</table>
</body>
</html>
<%end if%>
