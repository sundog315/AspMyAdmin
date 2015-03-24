<%if Session("admin")=true then%>
<html>
<head><title>sqlalter</title>
<script LANGUAGE="JavaScript">
	function validate(theForm) {
	if (theForm.type.value == "")
  	{
    alert("请输入数据类型");
    theForm.type.focus();
    return (false);
  	}
	return (true);
  }
</script>
<link rel="stylesheet" href="MASTER.CSS" type="text/css">
</head>
<body bgcolor="#DEF3FE" class="detaillabels">
<!--#include file="config.asp"-->
<!--#include file="adovbs.inc"-->
<%
Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr 

action=request("action")
column=request("column")

if column<>"" then
session("column")=column
end if

table=request("table")
if table<>"" then
	Session("table")=table
end if

if request("submit")<>"" then
   sql="ALTER+TABLE+["&session("table")&"]+"
   sql=sql&"ALTER+COLUMN+["&session("column")&"]+"
   if request("type")<>"" then
      sql=sql&request("type")
   end if
   if request("size")<>"" and request("null")="" then
      sql=sql&"("&request("size")&")+"
   end if
   if request("size")<>"" and request("null")<>"" then
      sql=sql&"("&request("size")&")"
   end if
   if request("null")<>"ON" then
      sql=sql&"NOT+NULL"
   end if
sql=sql&";"
response.redirect "sql.asp?submit_confirm=sundog&sql="&sql
end if

if action="alter_change" and request("submit")="" then
sql="SELECT * FROM ["&table&"]"
  set rs=con.execute(sql)
  for i=0 to rs.fields.count-1
  if rs(i).name=column then
%>
<center>
<form action="sql_alter.asp" method="post" name=frm onsubmit="return validate(frm)">
    <table width="75%" border="1" cellspacing="0" cellpadding="0">
      <tr> 
        <th class="detaillabels" width="13%"> 
          <div align="center">字段</div>
        </th>
        <th class="detaillabels" width="13%"> 
          <div align="center">类型</div>
        </th>
        <th class="detaillabels" width="20%"> 
          <div align="center">设定大小</div>
        </th>
        <th class="detaillabels" width="25%">允许空值</th>
        <th class="detaillabels" width="25%"> 
          <div align="center">自动编号</div>
        </th>
        <th class="detaillabels" width="16%"> 
          <div align="center"></div>
        </th>
      </tr>
      <tr> 
        <td class="detaillabels" width="13%"> 
          <div align="center"><%=rs(i).name%></div>
        </td>
        <td class="detaillabels" width="13%"> 
          <div align="center"> 
            <select name="type">
              <option value=""></option>
              <option value="binary">BINARY</option>
              <option value="bit">BIT(是/否)</option>
              <option value="char">CHAR</option>
              <option value="datetime">DATETIME</option>
              <option value="float">FLOAT</option>
              <option value="image">DOUBLE</option>
              <option value="int">INT</option>
              <option value="money">MONEY</option>
              <option value="numeric">NUMERIC</option>
              <option value="real">REAL</option>
              <option value="smallint">SMALLINT</option>
              <option value="memo">MEMO</option>
              <option value="varchar">VARCHAR</option>
            </select>
          </div>
        </td>
        <td class="detaillabels" width="20%"> 
          <div align="center"> 
            <input type="text" name="size" size="13">
          </div>
        </td>
        <td class="detaillabels" width="25%"> 
          <div align="center"> 
<%
  attrib=rs(i).Attributes
  if (attrib and adFldIsNullable)=0 then
%>
            <input type="checkbox" name="null" value="checkbox">
<%
  else
%>
            <input type="checkbox" name="null" value="checkbox" checked>
<%
  end if
%>
          </div>
        </td>
        <td class="detaillabels" width="25%"> 
          <div align="center"> 
            <input type="checkbox" name="autoincrement" value="checkbox">
          </div>
        </td>
        <td class="detaillabels" width="16%"> 
          <div align="center"> 
            <input type="submit" name="Submit" value="提交">
          </div>
        </td>
      </tr>
    </table>
</form>
</center>
<%end if
  next
  end if
%>
</body>
</html>
<%end if%>
