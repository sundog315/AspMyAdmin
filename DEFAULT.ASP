<%if Request("Submit")<>"" then%>
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<%
session("admin")=false
session("user_name")=""
Set Con=Server.CreateObject("ADODB.Connection")
Con.Open constr
set objSchema = Con.OpenSchema(20,Array(Empty, Empty, Empty, "TABLE"))        
Do until objSchema.EOF        
	if objSchema("TABLE_NAME")="AspMyAdmin_user" then
		sql="Select * From [AspMyAdmin_user] WHERE [user]='" & Request("user_name") & "' AND [password]='" & Base64encode(Request("password")) & "'"
		Set rs=Con.Execute(sql)
		if rs.eof then
			Response.Redirect "error.htm"
		    objSchema.Close
		    Set objSchema=nothing
		    Con.Close
			Set Con=nothing
		else
			session("admin")=true
			session("user_name")=Request("user_name")
			Response.Redirect "index.asp"
			objSchema.Close
 		   Set objSchema=nothing
			Con.Close
			Set Con=nothing
		end if
	end if
    objSchema.MoveNext
Loop
sql="Create Table [AspMyAdmin_user] ([user] VarChar(50) IDENTITY PRIMARY KEY, [password] VarChar(50))"
Con.Execute(sql)
sql="Insert INTO [AspMyAdmin_user] VALUES ('" & Request("user_name") & "','" & Base64encode(Request("password")) & "')"
'response.write sql
'response.end
Con.Execute(sql)
session("admin")=true
session("user_name")=Request("user_name")
objSchema.Close
Set objSchema=nothing
Con.Close
Set Con=nothing
response.redirect "first.htm"
else
%>
<html>
<head>
<title>AspMyAdmin</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="master.css" type="text/css">
</head>

<body bgcolor="#DEF3FE" text="#000000">
<div align="center">
  <p><font face="方正姚体" size="5">AspMyAdmin数据库管理系统</font></p>
  <p>&nbsp;</p>
<form action="default.asp" method="post">
  <table width="241" border="1" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="104" class="form"> 
        <div align="center">用户名</div>
      </td>
      <td width="131" class="form"> 
        <div align="center"> 
            <input type="text" name="user_name" size="15">
        </div>
      </td>
    </tr>
    <tr> 
      <td width="104" class="form"> 
        <div align="center">密码</div>
      </td>
      <td width="131" class="form"> 
        <div align="center"> 
            <input type="password" name="password" size="15">
        </div>
      </td>
    </tr>
    <tr> 
      <td colspan="2" class="form"> 
        <div align="center"></div>
        <div align="center"> 
          <input type="submit" name="Submit" value="提交">
          <input type="reset" name="ReSet" value="重设">
        </div>
      </td>
    </tr>
  </table>
</form>
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
  <p>&nbsp;</p>
</div>
</body>
</html>
<%end if%>
