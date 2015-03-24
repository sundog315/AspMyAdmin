<%if session("admin")=true then%>
<!--#include file="adovbs.inc"-->
<!--#include file="config.asp"-->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>导出</title>
<link rel="stylesheet" href="MASTER.CSS" type="text/css">
</head>

<body bgcolor="#DEF3FE">
<%
ON ERROR RESUME NEXT
if request("struct")="ON" then%>

<p class="hei"><span class="detaillabels">#以下是表<%=request("table")%>的结构:</span></p>
<span class="detaillabels">-------------------------------------------------------- <br>
<%
table="["&request("table")&"]"
Session("table")=table

Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr 

Set primary = Con.OpenSchema(adSchemaPrimaryKeys, _
        Array(empty, empty, table))

if primary("COLUMN_NAME")<>"" then
	primarykey=primary("COLUMN_NAME")
end if

primary.Close
set primary=nothing

tbl_struct="CREATE TABLE "&table&" ("
sql="SELECT * FROM "&table
	Set rs=Con.Execute(sql)
for i=0 to rs.fields.count-1
	tbl_struct=tbl_struct&"["&rs(i).name&"] "
	 field_type=rs(i).type

Select Case field_type
   case 3 typ="INT"
   case 2 typ="SMALLINT"
   case 4 typ="REAL"
   case 6 typ="MONEY"
   case 5 typ="NUMERIC"
   case 7 typ="DATETIME"
   case 11 typ="BIT"
   case 130 typ="CHAR"
   case 202 typ="VARCHAR"
   case 203 typ="MEMO"
   case 205 typ="IMAGE"
   case 128 typ="BINARY"
End Select

   if typ="VARCHAR" or typ="BINARY" or typ="CHAR" then
     tbl_struct=tbl_struct&typ&"("&rs(i).definedsize&")"
   else
     tbl_struct=tbl_struct&typ&" "
   end if

   attrib=rs(i).attributes

   if (attrib and adFldIsNullable)=0 then
     tbl_struct=tbl_struct&" NOT NULL"
   end if

   if rs(i).Properties("ISAUTOINCREMENT") = True then
     tbl_struct=tbl_struct&" IDENTITY"
   end if

     tbl_struct=tbl_struct&","
next

if primarykey<>"" then
  tbl_struct=tbl_struct&"PRIMARY KEY (["&primarykey&"]))"
else
  len_of_sql=Len(tbl_struct)
  tbl_struct=Mid(tbl_struct,1,len_of_sql-1)
  tbl_struct=tbl_struct&")"
end if
%>
<%=tbl_struct%> 
<%
rs.close
set rs=nothing
%>
<%
end if

if request("data")="ON" then%>
</span>
<p class="hei"><span class="detaillabels">#以下是表<%=request("table")%>的数据:</span></p>
<p class="sundog"><span class="detaillabels"> --------------------------------------------------------</span></p>
<p> <span class="detaillabels">
  <%
  sql="SELECT * FROM "&table
    Set rs=Con.Execute(sql)
    while not rs.eof
    str="INSERT INTO "&table&" VALUES ("
    for i=0 to rs.fields.count-1               
      parsevalue=rs(i).value
       select case rs(i).type
          case adDBDate
            str=str&"#"&parsevalue&"#,"
          case adDBTime
            str=str&"#"&parsevalue&"#,"
          case adDBTimeStamp
            str=str&"#"&parsevalue&"#,"
          case adBSTR
            str=str&"'"&server.htmlencode(parsevalue)&"',"
          case adChar
            str=str&"'"&server.htmlencode(parsevalue)&"',"
          case adVarChar
            str=str&"'"&server.htmlencode(parsevalue)&"',"
          case adLongVarChar
            str=str&"'"&server.htmlencode(parsevalue)&"',"
          case adWChar
            str=str&"'"&server.htmlencode(parsevalue)&"',"
          case adVarWChar
            str=str&"'"&server.htmlencode(parsevalue)&"',"
          case adLongVarWChar
            str=str&"'"&server.htmlencode(parsevalue)&"',"
          case else
            str=str&server.htmlencode(parsevalue)&","
  end select
     next
  len_of_sql=Len(str)
  str=Mid(str,1,len_of_sql-1)
  str=str&")"
%>
  </span></p>
<p class="sundog"> <span class="detaillabels"><%=str%></span></p>
<%
  rs.movenext  
  wend
  rs.close
  set rs=nothing
  end if
  con.close
  set con=nothing%>
<%end if%>
