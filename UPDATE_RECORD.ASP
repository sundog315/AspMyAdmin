<%if Session("admin")=true then%>
<!--#include file="adovbs.inc"-->
<!--#include file="config.asp"-->
<%
Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr

col_value=Request("col_value")
coloum=Request("coloum")

if Request("submit_confirm")<>"" then
  sql_str="SELECT * FROM [" & Session("table") & "] WHERE [" & coloum & "]=" & col_value
  set rs=Con.execute(sql_str)

     update_sql="UPDATE [" & Session("table") & "] SET"
     j=0
     for i=0 to rs.fields.count-1
       if Request(rs(i).Name)<>"" then
            update_sql=update_sql & " [" & rs(i).Name & "]"
          if rs(i).Type=3 then
            update_sql=update_sql & "=" & Request(rs(i).Name) & ","
          elseif rs(i).Type=7 then
            update_sql=update_sql & "=" & "#" & Request(rs(i).Name) & "#,"
          else
            update_sql=update_sql & "=" & "'" & Request(rs(i).Name) & "',"
          end if          
       end if
     next

len_of_sql=Len(update_sql)
update_sql=Mid(update_sql,1,len_of_sql-1)

update_sql=update_sql & " WHERE [" & coloum & "]"
update_sql=update_sql & "=" & col_value

rs.close
set rs=nothing

Con.Execute(update_sql)
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
<%
else
%>
<html>
<head>
<title>更新一条记录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="master.css" type="text/css">
</head>

<body bgcolor="#DEF3FE">
<div align="center">
<form action=update_record.asp method=post>
    <table width="75%" border="1" cellpadding="0" cellspacing="0">
      <tr> 
        <%
Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr

coloum=Request("coloum")
col_value=Request("col_value")
sql="SELECT * FROM [" & session("table") & "] WHERE [" & coloum & "]=" & col_value
set rs=Con.execute(sql)

for i=0 to rs.fields.count-1
%>
        <th class="detaillabels"> 
          <div align="center"><%=rs(i).Name%></div>
        </th>
        <%next %>
        <th class="detaillabels"> 
          <div align="center"></div>
        </th>
      </tr>
      <tr> 
        <%for i=0 to rs.fields.count-1%>
        <td class="detaillabels"> 
          <div align="center"> <%=rs(i).Value%> </div>
        </td>
        <%next%>
        <td class="detaillabels">&nbsp;</td>
      </tr>
      <tr> 
        <%for i=0 to rs.fields.count-1%>
        <td class="detaillabels"> 
          <div align="center"> 
            <input type=text name=<%=rs(i).name%> size="10">
          </div>
        </td>
        <%
  next
  rs.close
  set rs=nothing
  con.close
  set con=nothing
%>
        <td class="detaillabels"> 
          <div align="center"> 
            <input type=hidden name=col_value value=<%=col_value%>>
            <input type=hidden name=coloum value=<%=coloum%>>
            <input type=submit name=submit_confirm value="提交">
          </div>
        </td>
      </tr>
    </table>
</form>
  <p>&nbsp;</p>
</div>
</body>
</html>
<%end if
end if%>
