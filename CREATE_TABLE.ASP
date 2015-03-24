<%if session("admin")=true then%>
<%
on error resume next
field_num=Request("field_num")
table_name=request("table_name")
%>
<html>
<head>
<title>创建新表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="MASTER.CSS" type="text/css">
</head>

<body bgcolor="#DEF3FE">
<%if request("submit")="" then%>
<center>
  <p class="hei"><span class="detaillabels">创建表：</span><%=table_name%></p>
  <form action="CREATE_TABLE.ASP" method="post">
    <table width="510" border="1" cellspacing="0" cellpadding="0">
      <tr> 
        <td class="detaillabels" align="center" width="78"> 
          <div align="center">字段名</div>
        </td>
        <td class="detaillabels" align="center" width="92"> 
          <div align="center">类型</div>
        </td>
        <td class="detaillabels" align="center" width="78"> 
          <div align="center">大小</div>
        </td>
        <td class="detaillabels" align="center" width="92"> 
          <div align="center">空值</div>
        </td>
        <td class="detaillabels" align="center" width="51"> 自动编号 </td>
        <td class="detaillabels" align="center" width="105"> 
          <div align="center">主键</div>
        </td>
      </tr>
      <%for i=0 to field_num-1%>
      <tr> 
        <td class="detaillabels" align="center" width="78"> 
          <div align="center"> 
            <input type="text" name="field_name" size="10">
          </div>
        </td>
        <td class="detaillabels" align="center" width="92"> 
          <div align="center"> 
            <select name="field_type">
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
        <td class="detaillabels" align="center" width="78"> 
          <div align="center"> 
            <input type="text" name="field_size" size="10">
          </div>
        </td>
        <td class="detaillabels" align="center" width="92"> 
          <div align="center"> 
            <select name="null">
              <option value="NOT_NULL">NOT_NULL</option>
              <option value="NULL">NULL</option>
            </select>
          </div>
        </td>
        <td class="detaillabels" align="center" width="51"> 
          <select size="1" name="autoincrement">
            <option></option>
            <option>自动编号</option>
          </select>
          　 </td>
        <td class="detaillabels" align="center" width="105"> 
          <div align="center"> 
            <select name="primarykey">
              <option></option>
              <option value="primarykey">primarykey</option>
            </select>
          </div>
        </td>
      </tr>
      <%next%>
      <tr> 
        <td class="detaillabels" colspan="5" width="399"> 
          <div align="right"> 
            <input type="hidden" name="i" value=<%=i%>>
            <input type="hidden" name="table_name" value="<%=table_name%>">
            <input type="submit" name="Submit" value="提交">
          </div>
        </td>
		<td class="detaillabels">&nbsp;
				
		</td>
      </tr>
    </table>
</form>
</center>
<%else
sql="CREATE TABLE ["&table_name&"] ("
for i=1 to request("i")
   sql=sql&"["&request("field_name")(i)&"] "&request("field_type")(i)
      if request("field_size")(i)<>"" then
          sql=sql&"("&request("field_size")(i)&")"
      end if
      if request("null")(i)="NOT_NULL" then
          sql=sql&" not null"
      end if
      if request("autoincrement")(i)="自动编号" then
          sql=sql&" identity"
      end if
      if request("primarykey")(i)="primarykey" then
          primarykey=request("field_name")(i)
      end if
if primarykey<>"" then
   sql=sql&","
end if
next
if primarykey<>"" then
   sql=sql&" primary key (["&primarykey&"]) "
end if
sql=sql&");"
response.redirect "sql.asp?submit_confirm=sundog&sql="&sql 
end if
%>
</body>
</html>
<%end if%>
