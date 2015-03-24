<%if Session("admin")=true then%>
<%response.buffer=true%>
<!--#include file="adovbs.inc"-->
<!--#include file="config.asp"-->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>运行SQL语句</title>
<link rel="stylesheet" href="MASTER.CSS" type="text/css">
</head>

<body bgcolor="#DEF3FE" class="detaillabels">
<%
On Error Resume Next

Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr

limit=20

action=Request("action")
if not isempty(action) then
	Session("action")=action
end if

sql=Request("sql")
if not isempty(sql) then
	Session("sql")=sql
end if

if instr(sql,"select")>0 then
	Session("action")="select"
end if

table=request("table")
if not isempty(table) then
	Session("table")=table
end if


if Request("submit_confirm")<>"" and Session("action")<>"select"then
sql=trim(Request("sql"))
Set rs=Server.CreateObject("ADODB.RecordSet")
CursorLocation=adUseClient
sqlprased=Split(sql,";",-1,1)
i=0
DO
  if isempty(sqlprased(i)) then
    exit do
  end if
  sqlstr=sqlprased(i)
  rs.Open sqlstr, Con, 3, 3
  i=i+1
Loop
For Each errLoop In con.Errors
	Response.Write "发生错误,以下是具体信息:<br><br>"
	Response.Write errLoop.Number&"<br>"
	Response.Write errLoop.Description&"<br>"
	Response.Write "Source: "&errLoop.Source&"<br>"
	Response.Write "SQL State: "&errLoop.SQLState&"<br>"
	Response.Write "NativeError: "&errLoop.NativeError&"<br>"
	Con.Errors.Clear
	Session("action")=""
	Session("sql")=""
	Session("table")=""
	Response.Redirect "main.asp"
Next
Con.Close
Set Con=Nothing
end if


if Session("action")="delete" or Session("action")="drop" then%>
<center>
<form action=sql.asp?sql=<%=server.urlencode(sql)%> method=post>
    <table width="79%" border="0" cellpadding="0" cellspacing="0">
      <tr> 
        <td class="detaillabels"> 
          <h2 align="center">您选择了<%=Session("action")%>操作</h2>
        </td>
      </tr>
      <tr> 
        <td class="detaillabels">SQL:=<font color=red><%=Session("sql")%></font></td>
      </tr>
      <tr> 
        <td class="detaillabels" height="9"> 
          <div align="center"> 
            <input type="submit" name="Submit_confirm" value="提交">
          </div>
        </td>
      </tr>
    </table>
</form>
</center>
<p> 
  <%
end if


if Session("action")="select" then
	Set rs=Server.CreateObject("ADODB.RecordSet")
	rs.Open Session("sql"), con, 3, 3 
	if rs.RecordCount=0 then
		Response.Redirect "main.asp"
	end if
	rs.PageSize=30
	page=Request("page")
	if isempty(page) then
		page=1
	end if
	rs.AbsolutePage=page %>
</p>
<table width=100% border=1 cellpadding=0 cellspacing=0>
  <tr> 
    <%for i=0 to rs.fields.count-1%>
    <th class="detaillabels" width="560"><%=rs(i).name%></th>
    <%next %>
    <th class="detaillabels" width="84">操作</th>
  </tr>
  <%
if not rs.eof then
  j=1  
while not rs.eof and j<rs.pagesize%>
  <tr> 
    <%for i=0 to rs.fields.count-1 
   val=rs(i).value 
if isempty(val) then
   val=" "
end if
%>
    <td class="detaillabels" width="560"><%=Server.HtmlEncode(val)%></td>
    <%next %>
    <td class="detaillabels" width="84"> 
      <table width="93%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td class="detaillabels"> 
            <div align="center"> 
              <%
Set primary = Con.OpenSchema(adSchemaPrimaryKeys, _
        Array(empty, empty,Session("table")))
if primary("COLUMN_NAME")<>"" then
primarykey=primary("COLUMN_NAME")
end if
primary.Close
Set primary=Nothing

del_sql="Delete From [" & Session("table") & "] WHERE "
for i=0 to rs.Fields.count-1
	if Not IsEmpty(primarykey) then
		if rs(i).Name = primarykey then
			if rs(i).Type = adInteger then
				del_sql = del_sql & rs(i).Name & "=" & rs(i).Value
				primarykey_value = rs(i).Value
			elseif rs(i).Type = adDate then
				del_sql = del_sql & rs(i).Name & "=#" & rs(i).Value & "#"
				primarykey_value = "#" & rs(i).Value & "#"
			else
				del_sql = del_sql & rs(i).Name & "='" & rs(i).Value & "'"
				primarykey_value = "'" & rs(i).Value & "'"
			end if
		end if
	else
		if rs(i).Type = adInteger then
			del_sql = del_sql & " " & rs(i).Name & " = " & rs(i).Value & " AND"
		elseif rs(i).Type = adDate then
			del_sql = del_sql & " " & rs(i).Name & " = #" & rs(i).Value & "# AND"
		else
			del_sql = del_sql & " " & rs(i).Name & " = '" & rs(i).Value & "' AND"
		end if
	end if
next

if IsEmpty(primarykey) then
	len_of_sql = Len(del_sql)
	del_sql=Mid(del_sql,1,len_of_sql-4)
end  if

del_sql=del_sql & ";"

'Response.Write del_sql & "<BR>"
'Response.End

del_sql=Server.UrlEncode(del_sql)
%>
              <a href=sql.asp?action=delete&sql=<%=del_sql%>>删除</a> </div>
          </td>
          <td class="detaillabels"> 
            <div align="center"><a href=update_table.asp?table=<%=session("table")%>&coloum=<%=primarykey%>&col_value=<%=primarykey_value%>>修改</a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <%    
j=j+1
rs.MoveNext     
wend
end if%>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="detaillabels"> 
      <div align="right"> 
        <%if page-1 >0 then
          response.write "<a href=sql.asp?page="&cstr(page-1)&">"
          response.write "上一页"
          response.write "</a>"
          response.write "&nbsp;"
          end if
          if (page-limit)>0 then
          response.write "<a href=sql.asp?page="&cstr(page-limit)&"><<<</a>&nbsp;"
          end if
          if (limit-page)>0 and page<>1 then
          response.write "<a href=sql.asp?page=1><<<</a>&nbsp;"
          end if
          for n=page to page+limit
            if n>rs.pagecount then
               exit for
            end if
          response.write "<a href=sql.asp?page="&cstr(n)&">"&cstr(n)&"</a>&nbsp;"
          next
          if (page+limit)<rs.pagecount then
          response.write "<a href=sql.asp?page="&cstr(page+limit)&">>>></a>&nbsp;"
          end if
          if page+1<=rs.pagecount then
          response.write "<a href=sql.asp?page="&cstr(page+1)&">"
          response.write "下一页"
          response.write "</a>"
          response.write "&nbsp;"
          end if
          %>
      </div>
    </td>
  </tr>
</table>
  <%
rs.close
set rs=nothing
con.close
set con=nothing
end if
%>
</body>
</html>
<%end if%>
