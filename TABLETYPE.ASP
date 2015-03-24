<%if Session("admin")=true then%>
<%response.buffer=true%>
<!--#include file="adovbs.inc"-->
<!--#include file="config.asp"-->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>表类型</title>
<link rel="stylesheet" href="MASTER.CSS" type="text/css">
</head>

<body bgcolor="#DEF3FE">
<%
'ON ERROR RESUME NEXT

Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr 

table=Request("table")
if table<>"" then
	Session("table")=table
Set primary = Con.OpenSchema(adSchemaPrimaryKeys, _
        Array(empty, empty, table))
if primary("COLUMN_NAME")<>"" then
primarykey=primary("COLUMN_NAME")
end if
primary.Close
Set primary=Nothing
%>
<center>
  <table width="700" border="1" cellpadding="0" cellspacing="0">
    <tr> 
      <th class="detaillabels" width="60"> 
        <div align="center">字段</div>
      </th>
      <th class="detaillabels" width="130"> 
        <div align="center">类型</div>
      </th>
      <th class="detaillabels" width="104"> 
        <div align="center">设定大小</div>
      </th>
      <th class="detaillabels" width="68"> 
        <div align="center">允许空值</div>
      </th>
      <th class="detaillabels" width="68"> 自动编号 </th>
      <th class="detaillabels" width="80">主键</th>
      <th class="detaillabels" width="153"> 
        <div align="center">执行操作</div>
      </th>
    </tr>
    <%sql="SELECT * FROM ["&table&"] "
	Set rs=Con.Execute(sql)
	For i=0 to rs.fields.count-1
%>
    <tr> 
      <td class="detaillabels" height="2" width="60"> 
        <div align="center"><%=rs(i).name%></div>
      </td>
      <td class="detaillabels" height="2" width="130"> 
        <div align="center"> 
          <%
field_type=rs(i).type
Select Case field_type
    case adEmpty
      typ = "Empty"
    case adTinyInt
      typ = "TinyInt"
    case adSmallInt
      typ = "SmallInt"
    case adInteger
      typ = "Integer"
    case adBigInt
      typ = "BigInt"
    case adUnsignedTinyInt
      typ = "UnsignedTinyInt"
    case adUnsignedSmallInt
      typ = "UnsignedSmallInt"
    case adUnsignedInt
      typ = "UnsignedInt"
    case adUnsignedBigInt
      typ = "UnsignedBigInt"
    case adSingle
      typ = "Single"
    case adDouble
      typ = "Double"
    case adCurrency
      typ = "Currency"
    case adDecimal
      typ = "Decimal"
    case adNumeric
      typ = "Numeric"
    case adBoolean
      typ = "Boolean"
    case adError
      typ = "Error"
    case adUserDefined
      typ = "UserDefined"
    case adVariant
      typ = "Variant"
    case adIDispatch
      typ = "IDispatch"
    case adIUnknown
      typ = "IUnknown"
    case adGUID
      typ = "GUID"
    case adDATE
      typ = "DATE"
    case adDBDate
      typ = "DBDate"
    case adDBTime
      typ = "DBTime"
    case adDBTimeStamp
      typ = "DBTimeStamp"
    case adBSTR
      typ = "BSTR"
    case adChar
      typ = "Char"
    case adVarChar
      typ = "VarChar"
    case adLongVarChar
      typ = "LongVarChar"
    case adWChar
      typ = "WChar"
    case adVarWChar
      typ = "VarWChar"
    case adLongVarWChar
      typ = "LongVarWChar"
    case adBinary
      typ = "Binary"
    case adVarBinary
      typ = "VarBinary"
    case adLongVarBinary
      typ = "LongVarBinary"
    case adChapter
      typ = "Chapter"
    case adPropVariant
      typ = "PropVariant"
    case else
      typ = "Unknown"
  end select
Response.Write typ%>
        </div>
      </td>
      <td class="detaillabels" height="2" width="104"> 
        <div align="center"><%=rs(i).definedsize%></div>
      </td>
      <td class="detaillabels" height="2" width="68"> 
        <div align="center"> 
          <%
  attrib=rs(i).Attributes
  if (attrib and adFldIsNullable)=0 then
	Response.Write "No"
  else
	Response.Write "Yes"
  end if
%>
        </div>
      </td>
      <td class="detaillabels" height="2" width="68"> 
        <div align="center"> 
          <%if rs(i).Properties("ISAUTOINCREMENT") = True then%>
          <input type="checkbox" name="autoincrement" value="checkbox" checked>
          <%else%>
          <input type="checkbox" name="autoincrement" value="checkbox">
          <%end if%>
        </div>
      </td>
      <td class="detaillabels" height="2" width="80"> 
        <div align="center"> 
          <%if rs(i).name=primarykey then%>
          <input type="checkbox" name="primarykey" value="checkbox" checked>
          <%else%>
          <input type="checkbox" name="primarykey" value="checkbox">
          <%end if%>
        </div>
      </td>
      <td class="detaillabels" height="2" width="153"> 
        <div align="center"> 
          <table width="86%" border="0" cellpadding="0" cellspacing="0" height="8">
            <tr> 
              <td class="detaillabels"> 
                <div align="center"><a href="SQL_ALTER.ASP?action=alter_change&Column=<%=rs(i).name%>&table=<%=table%>">改变</a></div>
              </td>
              <td class="detaillabels"> 
                <div align="center"><a href="sql.asp?action=drop&sql=alter+table+[<%=table%>]+drop+[<%=rs(i).name%>];">丢弃</a></div>
              </td>
              <td class="detaillabels"> 
                <p align="center"><a href="sql.asp?submit_confirm=sundog&sql=alter+table+[<%=table%>]+add+primary+key+([<%=rs(i).name%>]);">主键</a></p>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
    <%next %>
  </table>
</center>
<p align="left" class="sundog"><span class="detaillabels">增加字段：</span></p>
 
<form action="ALTER_ADD.ASP" method="post">
  <table border="1" cellpadding="0" cellspacing="0" width="443" height="39">
    <tr> 
      <td height="12" class="detaillabels" width="60" align="center">字段</td>
      <td height="12" class="detaillabels" width="50" align="center">类型</td>
      <td height="12" class="detaillabels" width="58" align="center">设定大小</td>
      <td height="12" class="detaillabels" width="64" align="center">允许空值</td>
      <td height="12" class="detaillabels" width="66" align="center"> 自动编号</td>
      <td height="12" class="detaillabels" width="96" align="center">&nbsp;&nbsp;</td>
    </tr>
    <tr> 
      <td height="23" class="detaillabels" width="60" align="center"> 
        <input type="text" name="fldname" size="10">
      </td>
      <td height="23" class="detaillabels" width="50" align="center"> 
        <select size="1" name="fldtype">
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
      </td>
      <td height="23" class="detaillabels" width="58" align="center"> 
        <input type="text" name="fldsize" size="10">
      </td>
      <td height="23" class="detaillabels" width="64" align="center"> 
        <input type="checkbox" name="null" value="ON">
      </td>
      <td height="23" class="detaillabels" width="66" align="center"> 
        <input type="checkbox" name="autoincrement" value="ON">
      </td>
      <td height="23" class="detaillabels" width="96" align="center"> 
        <input type="hidden" name="table" value="[<%=table%>]">
        <input type="submit" value="提交">
      </td>
    </tr>
  </table>
</form>
  
<p align="left" class="detaillabels">添加一条数据：</p>
<form action=insert_record.asp?Submit=sundog method=post>
  <table width="75%" border="1" cellpadding="0" cellspacing="0">
    <tr> 
      <%
for i=0 to rs.fields.count-1%>
      <th class="detaillabels"><%=rs(i).name%></th>
      <%next%>
      <td class="detaillabels">&nbsp;</td>
    </tr>
    <tr> 
      <%for i=0 to rs.fields.count-1%>
      <td class="detaillabels"> 
        <input type=text name=<%=rs(i).name%> size="10">
      </td>
      <%next%>
      <td class="detaillabels"> 
        <input type=hidden name=table value=<%=table%>>
        <input type=submit name=submit_confirm value="提交">
      </td>
    </tr>
  </table>
</form>
<p align="left" class="detaillabels">导出--SQL</p>
<form action="DB2SQL.ASP" method="post">
  <table border="1" cellpadding="0" cellspacing="0" width="187" height="39">
    <tr> 
      <td width="90" align="center" class="detaillabels">数据</td>
      <td width="117" align="center" class="detaillabels">结构</td>
      <td width="132" align="center" class="detaillabels">&nbsp;</td>
    </tr>
    <tr> 
      <td width="90" align="center" class="detaillabels"> 
        <input type="checkbox" name="data" value="ON">
      </td>
      <td width="117" align="center" class="detaillabels"> 
        <input type="checkbox" name="struct" value="ON">
      </td>
      <td width="132" align="center" class="detaillabels"> 
        <input type="hidden" name="table" value="<%=table%>">
        <input type="submit" value="提交">
      </td>
    </tr>
  </table>
</form>   
<%        
Con.Close        
Set Con=Nothing 
end if
%>
</body>
</html>
<%end if%>
