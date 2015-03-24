<%if session("admin")=true then%>
<%
fldname=request("fldname")
fldtype=request("fldtype")
fldsize=request("fldsize")
fldnull=request("null")
fldautoincrement=request("autoincrement")
table=request("table")
if fldname<>"" and fldtype<>"" then
  sql="alter+table+"&table&"+add+["&fldname&"]+"&fldtype
  
  if fldsize<>"" then
    sql=sql&"("&fldsize&")"
  end if 
  
  if fldnull<>"ON" then
    sql=sql&"+not+null"
  end if
  
  if fldautoincrement="ON" then
    sql=sql&"+indentity;"
  else
    sql=sql&";"
  end if
  response.redirect "sql.asp?submit_confirm=sundog&sql="&sql
else
  response.write "输入数据错误！"
end if
%>
<%end if%>
