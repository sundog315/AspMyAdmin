<%if session("admin")=true then%>
<!--#include file="adovbs.inc"-->
<!--#include file="config.asp"-->
<%
view_num=Request("view_num")
view_name=request("view_name")
Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr 
%>
<html>
<head>
<title>创建新视图</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="MASTER.CSS" type="text/css">
</head>
<body bgcolor="#DEF3FE" class="detaillabels">
<%if request("Submit_view")="" then%>
<center>
  <p class="hei">创建简单视图：<%=view_name%></p>
  <form action="CREATE_VIEW.ASP" method="post">
    <table width="510" border="1" cellspacing="0" cellpadding="0" class="detaillabels">
      <tr> 
        <td class="sundog" align="center" width="78"> 
          <div align="center">字段名</div>
        </td>
        <td class="sundog" align="center" width="92"> 
          <div align="center">表名</div>
        </td>
        <td class="sundog" align="center" width="78"> 
          <div align="center">字段名</div>
        </td>
      </tr>
      <%for i=0 to view_num-1%>
      <tr> 
        <td class="sundog" align="center" width="78"> 
          <div align="center"> 
            <input type="text" name="view_field" size="10">
          </div>
        </td>
        <td class="sundog" align="center" width="92"> 
          <div align="center"> 
            <select name="table_name">
              <%        
set objSchema = Con.OpenSchema(20,Array(Empty, Empty, Empty, "TABLE"))        
Do until objSchema.EOF        
strname = objSchema("TABLE_NAME")        
%>
              <option><%=strname%></option>
              <%objSchema.MoveNext
  loop
  objSchema.close
  set objSchema=nothing%>
            </select>
          </div>
        </td>
        <td class="sundog" align="center" width="78"> 
          <div align="center"> 
            <select name="field_name">
              <%        
set objSchema = Con.OpenSchema(20,Array(Empty, Empty, Empty, "TABLE"))        
Do until objSchema.EOF        
strname = objSchema("TABLE_NAME")  
set rs=server.createobject("adodb.recordset")
   rs.open "select * from ["&strname&"]", con, 3, 3
    for j=0 to rs.fields.count-1      
%>
              <option><%=strname&"."&rs(j).name%></option>
              <%  next
   rs.close
   set rs=nothing
  objSchema.MoveNext
  loop
  objSchema.close
  set objSchema=nothing%>
            </select>
          </div>
        </td>
      </tr>
      <%next%>
      <tr> 
        <td class="sundog" colspan="3"> 
          <div align="right">
            <input type="hidden" name="i" value="<%=i%>">
            <input type="hidden" name="view_name" value="<%=view_name%>">
            <input type="submit" name="Submit_view" value="提交">
          </div>
        </td>
      </tr>
    </table>
</form>
</center>
<%else
sql="CREATE VIEW ["&request("view_name")&"] ("
for i=1 to request("i")
   sql=sql&"["&request("view_field")(i)&"], "
next
len_of_sql=Len(sql)
sql=Mid(sql,1,len_of_sql-2)
sql=sql&") AS SELECT "
for i=1 to request("i")
   sql=sql&"["&request("field_name")(i)&"], "
next
len_of_sql=Len(sql)
sql=Mid(sql,1,len_of_sql-2)
sql=sql&" FROM "
for i=1 to request("i")
   if instr(str_from,request("table_name")(i)) then
   else
      str_from=str_from&"["&request("table_name")(i)&"], "
   end if
next
len_of_str_from=Len(str_from)
str_from=Mid(str_from,1,len_of_str_from-2)
sql=sql&str_from&";"
response.redirect "sql.asp?submit_confirm=sundog&sql="&sql 
end if
%>
</body>
</html>
<%end if%>
