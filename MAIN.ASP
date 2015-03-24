<%if Session("admin")=true then%>
<!--#include file="config.asp"-->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>工作区</title>
<link rel="stylesheet" href="MASTER.CSS" type="text/css">
<script language="javascript">
this.parent.contents.location.reload();
</script>
</head>

<body bgcolor="#DEF3FE">
<center>
  <table width="58%" border="1" cellpadding="0" cellspacing="0">
    <tr> 
      <th height="19" width="24%" class="detaillabels"> 
        <div align="center">表名</div>
      </th>
      <th height="19" width="46%" class="detaillabels"> 
        <div align="center">操作</div>
      </th>
      <th height="19" width="30%" class="detaillabels"> 
        <div align="center">记录数</div>
      </th>
    </tr>
    <% 
Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr 
 
Set objSchema = Con.OpenSchema(20,Array(Empty, Empty, Empty, "TABLE"))        
Do until objSchema.EOF        
	strName = objSchema("TABLE_NAME")        
%>
    <tr> 
      <td width="24%" class="detaillabels"><A HREF="TABLETYPE.ASP?table=<%=strname%>"><%=strname%></A> 
      </td>
      <td width="46%" class="detaillabels"> <center> 
        <div align="center"> 
          <table width="86%" border="0" cellpadding="0" cellspacing="0" height="8">
            <tr> 
              <td class="detaillabels"> 
                <div align="center"> <a href="tableview.asp?table=<%=strname%>">浏览</a> 
                </div>
              </td>
              <td class="detaillabels"> 
                <div align="center"> 
                  <%        
str_sql="table="&strname&"&sql=drop+table+["&strname&"];&action=drop"        
%>
                  <a href="sql.asp?<%=str_sql%>">丢弃</a></div>
              </td>
              <td class="detaillabels"> 
                <p align="center"> 
                  <%        
str_sql="table="&strname&"&sql=delete+from+["&strname&"];&action=delete"        
%>
                  <a href="sql.asp?<%=str_sql%>">清空</a></p>
              </td>
            </tr>
          </table><center>
          </div>
      </td>
      <td width="30%" class="detaillabels"> 
        <%        
sql="SELECT * FROM ["&strname&"]"        
Set ll=Server.CreateObject("ADODB.RecordSet")        
ll.Open sql,Con,3        
records=ll.RecordCount        
Response.Write cstr(records)        
%>
      </td>
    </tr>
    <%        
ll.Close        
Set ll=Nothing        
    objSchema.MoveNext        
Loop        
objschema.Close        
Set objschema=Nothing        
Con.Close        
Set con=Nothing        
%>
  </table>
  <p class="detaillabels">执行sql语句：</p>        
<form action="sql.asp" method=post>        
    <table width="47%" border="1" cellpadding="0" cellspacing="0">
      <tr> 
        <td width="79%" height="74" class="detaillabels"> 
          <textarea name="sql" cols="40" rows="5"></textarea>
        </td>
        <td width="21%" height="74" class="detaillabels"> 
          <table width="98%" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td class="detaillabels">　</td>
            </tr>
            <tr> 
              <td class="detaillabels">每行以</td>
            </tr>
            <tr> 
              <td class="detaillabels">&quot;;&quot;结尾!</td>
            </tr>
            <tr> 
              <td class="detaillabels">　</td>
            </tr>
            <tr> 
              <td class="detaillabels"> 
                <div align="center"> 
                  <input type="submit" name="Submit_confirm" value="提交">
                </div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>        
</form>        
  <p class="detaillabels">创建新表：</p>        
<form action="create_table.asp" method=post>        
    <table width="33%" border="1" cellspacing="0" cellpadding="0">
      <tr> 
        <td class="detaillabels">表名：</td>
        <td class="detaillabels"> 
          <div align="center"> <span class="sundog"> 
            <input type="text" name="table_name" size="20">
            </span></div>
        </td>
      </tr>
      <tr> 
        <td class="detaillabels">字段数：</td>
        <td class="detaillabels"> 
          <div align="center"> <span class="sundog"> 
            <input type="text" name="field_num" size="20">
            </span></div>
        </td>
      </tr>
      <tr> 
        <td colspan="2" class="detaillabels"> 
          <div align="center"> 
            <input type="submit" name="Submit_create" value="提交">
          </div>
        </td>
      </tr>
    </table>        
</form>        
  <p class="detaillabels">创建新视图：</p>
<form action="create_view.asp" method="post">
    <table width="33%" border="1" cellpadding="0" cellspacing="0">
      <tr> 
        <td width="23%" class="detaillabels">视图名：</td>
        <td width="77%" class="detaillabels"><span class="sundog"> 
          <input type="text" name="view_name" size="20">
          </span></td>
      </tr>
      <tr> 
        <td width="23%" class="detaillabels">字段数：</td>
        <td width="77%" class="detaillabels"><span class="sundog"> 
          <input type="text" name="view_num" size="20">
          </span></td>
      </tr>
      <tr> 
        <td colspan="2" class="detaillabels"> 
          <div align="center"> 
            <input type="submit" name="Submit" value="提交">
          </div>
          <div align="center"></div>
        </td>
      </tr>
    </table>
</form>
</center>
</body>

</html>
<%end if%>
