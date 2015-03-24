<%if Session("admin")=true then%>
<!--#include file="config.asp"-->
<%
if Request("action")="delete" then
	sql=Request("sql")

	Set Con=Server.CreateObject("ADODB.Connection")
	Con.Open constr

	Con.Execute(sql)
%>
<html>
<head>
<meta HTTP-EQUIV="Refresh" content="2; URL=main.asp">
<title>操作成功</title>
<link rel="stylesheet" href="master.css" type="text/css">
</head>

<body bgcolor="#DEF3FE">

<p>&nbsp;</p>
<blockquote>
  <p class="detaillabels">操作成功执行</p>
  <p class="detaillabels">等待2秒回到主工作区或按<a href="main.asp">这里</a>.</p>
<center>
  <table width="265" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td class="detaillabels"> 
        <div align="center">----------------------------------------------------</div>
      </td>
    </tr>
    <tr> 
      <td class="detaillabels"> 
        <div align="center">本系统由</div>
      </td>
    </tr>
    <tr> 
      <td class="detaillabels"> 
        <div align="center">杭州电子工业学院信息管理与信息系统</div>
      </td>
    </tr>
    <tr> 
      <td class="detaillabels"> 
        <div align="center">刘磊制作</div>
      </td>
    </tr>
    <tr> 
      <td class="detaillabels"> 
        <div align="center">2001&copy; CopyRight</div>
      </td>
    </tr>
    <tr> 
      <td class="detaillabels"> 
        <div align="center">----------------------------------------------------</div>
      </td>
    </tr>
  </table>
</center>
</blockquote>
</body>
</html>
<%else%>
<title>存储过程</title>
<link rel="stylesheet" href="master.css" type="text/css">
<body bgcolor="#DEF3FE" class="detaillabels">
<table width="49%" border="1" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="26%" class="detaillabels">存储过程：</td>
    <td width="74%" class="detaillabels"> 
      <%
		Set Con=Server.CreateObject("ADODB.Connection")
		Con.Open constr

		sTableName = Request("table")
      	Set adox = Server.CreateObject("ADOX.Catalog")
		adox.ActiveConnection = Con
		Response.Write "<textarea cols=40 rows=5 class=""tbflat"">" & adox.Procedures(sTableName).Command.CommandText & "</textarea><br>"
%>
    </td>
  </tr>
</table>
<p><a href="showproc.asp?action=delete&sql=DROP+PROCEDURE+[<%=sTableName%>]">删除该存储过程？</a></p>
<%
  Set adox=Nothing
  Con.Close
  Set Con=Nothing
End if
end if
%>
