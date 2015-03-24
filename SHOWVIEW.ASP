<%if Session("admin")=true then%>
<!--#include file="config.asp"-->
<%
if Request("action")<>"" then
Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr

Con.Execute(Request("sql"))
Con.Close
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
<html>
<head>
<title>查看视图</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="master.css" type="text/css">
</head>

<body bgcolor="#DEF3FE">
<%
'ON ERROR RESUME NEXT

table=Request("table")
Session("table")=table
if not IsEmpty(table) then
	Session("table")=table
end if
sql=Request("sql")
Session("sql")=sql

Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr

Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open Session("sql"), Con, 3, 3
%>
<table width=100% border=1 cellpadding=0 cellspacing=0>
  <tr> 
    <%for i=0 to rs.fields.count-1%>
    <th class="detaillabels" width="560"><%=rs(i).name%></th>
    <%next %>
  </tr>
  <%
if not rs.eof then
  j=1  
while not rs.eof%>
  <tr> 
    <%for i=0 to rs.fields.count-1 
   val=rs(i).value 
%>
    <td class="detaillabels" width="560">
      <%if not IsNull(val) then%>
  		    <%=Server.HtmlEncode(val)%> 
      <%else%>
   		   &nbsp; 
      <%end if%>
      </td>
    <%next %>
  </tr>
  <%    
j=j+1
rs.MoveNext     
wend
end if%>
</table>
<p class="detaillabels">视图：</p>
<table width="49%" border="1" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="detaillabels" width="26%">原视图：</td>
    <td class="detaillabels" width="74%">
      <%
      	Set adox = Server.CreateObject("ADOX.Catalog")
		adox.ActiveConnection = Con
		Response.Write "<textarea cols=40 rows=5 class=""tbflat"">" & "CREATE VIEW [" & Session("table") & "] AS " & adox.Views(Session("table")).Command.CommandText & "</textarea><br>"
%>
      </td>
  </tr>
</table>
<span class="detaillabels"><a href="showview.asp?action=delete&sql=DROP VIEW [<%=Session("table")%>]">删除该视图？</a></span>
<p> <span class="detaillabels">
  <%
rs.Close
Set rs=Nothing
Con.Close
Set Con=Nothing
%>
</span></p>
</body>
</html>
<%end if
end if%>
