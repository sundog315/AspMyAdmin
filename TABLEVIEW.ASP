<%if Session("admin")=true then%>
<!--#include file="adovbs.inc"-->
<!--#include file="config.asp"-->
<%
if Request("Submit")<>"" then
	Set Con=Server.CreateObject("ADODB.Connection")
	Con.Open constr

'	Response.Write Request("Sql") & "<br>"
'	Response.End

	Con.Execute(Request("Sql"))
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
<%
Response.End
end if
%>
<%
	action=Request("action")
	if action="delete" then
%>
<html>
<head>
<title>Delete Record</title>
<link rel="stylesheet" href="master.css" type="text/css">
</head>
<body bgcolor="#DEF3FE" class="detaillabels">
<center>
<p>&nbsp;</p>
<form Method=Post Action="tableview.asp">
  <table width="75%" border="1" cellpadding="0" cellspacing="0">
    <tr> 
      <td class="detaillabels"> 
        <div align="center">您选择了<font color=red>删除</font>操作，请确认！</div>
      </td>
    </tr>
    <tr> 
      <td class="detaillabels"> 
        <div align="center"><%=Request("Sql")%></div>
      </td>
    </tr>
    <tr> 
      <td class="detaillabels"> 
        <div align="center"> 
          <input type="submit" name="Submit" value="确定">
			<input type="hidden" name="Sql" value="<%=Request("Sql")%>">
          <input type="reset" name="Cancel" value="取消">
        </div>
      </td>
    </tr>
  </table>
  <p>&nbsp;</p>
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
</form>
  <p>&nbsp;</p>
</center>
</body>
</html>
<%else%>
<html>
<head>
<title>TableView</title>
<link rel="stylesheet" href="master.css" type="text/css">
</head>
<body bgcolor="#DEF3FE" class="detaillabels">
<%
	On error resume Next
	Set Con=Server.CreateObject("ADODB.Connection")
	Con.Open constr

	limit=20
	table=Request("table")
	if not IsEmpty(table) then
		Session("table")=table
	end if
'	Response.Write table & "<BR>"
'	Response.End

	Set primary = Con.OpenSchema(adSchemaPrimaryKeys, _
    	    Array(empty, empty, Session("table")))
	if primary("COLUMN_NAME")<>"" then
		primarykey=primary("COLUMN_NAME")
	else
		primarykey=""
	end if
	primary.close
	set primary=nothing

	sql="SELECT * FROM [" & Session("table") & "]  "
	if not IsEmpty(primarykey) then
		sql=sql & "Order By [" & primarykey & "]"
	end if
	if not IsEmpty(sql) then
		Session("sql")=sql
	end if
	
'	Response.Write Session("sql") & "<BR>"
'	Response.End

	set rs=Server.CreateObject("ADODB.RecordSet")
	rs.Open Session("sql") , Con ,3 ,3
'	if rs.RecordCount=0 then
'		Response.Redirect "main.asp"
'	end if
	rs.PageSize=30
	page=Request("page")
	if IsEmpty(page) then
		page=1
	end if
	rs.AbsolutePage=page 
%>
<p class="detaillabels"><a href="insert_record.asp"> 添加一条记录</a></p>
<table width=100% border=1 cellpadding=0 cellspacing=0>
  <tr> 
    <%for i=0 to rs.fields.count-1%>
    <th class="detaillabels"> 
      <div align="center"><%=rs(i).name%></div>
    </th>
    <%next %>
    <th class="detaillabels"> 
      <div align="center">操作</div>
    </th>
  </tr>
  <%
if not rs.eof then
  j=1  
While not rs.eof and j<rs.PageSize%>
  <tr> 
    <%for i=0 to rs.fields.count-1 
   val=rs(i).value 
%>
    <td> 
      <div align="center"><span class="detaillabels"> 
        <%if not IsNull(val) then%>
        <%=Server.HtmlEncode(val)%> 
        <%else%>
        &nbsp; 
        <%end if%>
        </span></div>
    </td>
    <%next %>
    <td> 
        <table width="93%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td> 
              <div align="center"> <span class="detaillabels"> 
                <%
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

'Response.Write del_sql & "<BR>"
'Response.End

del_sql=Server.UrlEncode(del_sql)
%>
                <a href=tableview.asp?action=delete&sql=<%=del_sql%>>删除</a> 
                </span></div>
            </td>
            <td> 
			<%if Not IsEmpty(primarykey) then%>
              <div align="center" class="detaillabels"><a href=update_record.asp?coloum=<%=primarykey%>&col_value=<%=primarykey_value%>>修改</a></div>
			<%end if%>
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
    <td> 
      <div align="right"> <span class="detaillabels"> 
        <%if page-1 >0 then
          response.write "<a href=tableview.asp?page="&cstr(page-1)&">"
          response.write "上一页"
          response.write "</a>"
          response.write "&nbsp;"
          end if
          if (page-limit)>0 then
          response.write "<a href=tableview.asp?page="&cstr(page-limit)&"><<<</a>&nbsp;"
          end if
          if (limit-page)>0 and page<>1 then
          response.write "<a href=tableview.asp?page=1><<<</a>&nbsp;"
          end if
          for n=page to page+limit
            if n>rs.pagecount then
               exit for
            end if
          response.write "<a href=tableview.asp?page="&cstr(n)&">"&cstr(n)&"</a>&nbsp;"
          next
          if (page+limit)<rs.pagecount then
          response.write "<a href=tableview.asp?page="&cstr(page+limit)&">>>></a>&nbsp;"
          end if
          if page+1<=rs.pagecount then
          response.write "<a href=tableview.asp?page="&cstr(page+1)&">"
          response.write "下一页"
          response.write "</a>"
          response.write "&nbsp;"
          end if
          %>
        </span></div>
    </td>
  </tr>
</table>
<%
rs.close
set rs=nothing
con.close
set con=nothing
%>
</body>
</html>
<%end if
end if%>
