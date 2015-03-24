<%if Session("admin")=true then%>
<!--#include file="adovbs.inc"-->
<!--#include file="config.asp"-->
<%
if Request("Submit")<>"" then
	Set Con=Server.CreateObject("ADODB.Connection")
	Con.Open constr
	
	sql="SELECT * FROM [" & Session("table") & "]"
	Set rs=Server.CreateObject("ADODB.RecordSet")
	rs.Open sql, Con, 3, 3
	ins_sql="INSERT INTO [" & Session("table") & "] VALUES ("
	for i=0 to rs.fields.count-1
		if rs(i).Type = adInteger then
			ins_sql = ins_sql & Request(rs(i).Name) & ","
		elseif rs(i).Type = adDate then
			ins_sql = ins_sql & "#" & Request(rs(i).Name) & "#,"
		else
			ins_sql = ins_sql & "'" & Request(rs(i).Name) & "',"
		end if
	next

	len_of_sql = Len(ins_sql)
	ins_sql=Mid(ins_sql,1,len_of_sql-1)

	ins_sql=ins_sql & ")"

'	Response.Write ins_sql
'	Response.End

	Con.Execute(ins_sql)
	rs.Close
	Set rs=Nothing
	Con.Close
	Set Con=Nothing
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
<title>插入一条记录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="master.css" type="text/css">
</head>

<body bgcolor="#DEF3FE">
<%
	Set Con=Server.CreateObject("ADODB.Connection")
	Con.Open constr

	sql="SELECT * FROM [" & Session("table") & "]"
'	Response.Write sql
'	Response.End
	Set rs=Server.CreateObject("ADODB.RecordSet")
	rs.Open sql, con, 3, 3
%>
<center>
<form method="post" action="insert_record.asp">
  <table width="75%" border="1" cellpadding="0" cellspacing="0">
    <tr> 
      <%
	for i=0 to rs.fields.count-1
%>
      <th class="detaillabels"><%=rs(i).Name%></th>
<%
	next
%>
    </tr>
    <tr> 
<%for i=0 to rs.fields.count-1%>
      <td class="detaillabels"><input type=text name=<%=rs(i).Name%>></td>
<%next%>
    </tr>
    <tr> 
      <td class="detaillabels" colspan=<%=rs.fields.count%>>
        <div align="center">
          <input type="submit" name="Submit" value="提交">
          <input type="reset" name="Reset" value="重设">
        </div>
      </td>
    </tr>
  </table>
</form>
</center>
<p align="center">&nbsp;</p>
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
</body>
</html>
<%end if
end if%>
