<%if session("admin")=true then%>
<!--#include file="config.asp"-->

<%if Request("action")="refresh" then%>
  <script language="javascript">
    this.parent.location.reload();
  </script>
<%end if%>

<%
if Request("Submit")<>"" then
	Set Con=Server.CreateObject("ADODB.Connection")
	Con.Open constr
	sql="UPDATE [AspMyAdmin_user] SET "
	if NOT isnull(Request("c_pwd")) then
			sql=sql & "[password]='" & Request("c_pwd") & "' "
	end if
	sql=sql & "WHERE [user]='" & session("user_name") &"'"
	Con.Execute(sql)
	Con.Close
	Set Con=nothing
	Session("admin")=false
	Session("user_name")=""
%>
<html>
<head>
<title>更改密码</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta HTTP-EQUIV="Refresh" content="2; URL=chgpwd.asp?action=refresh">
<link rel="stylesheet" href="master.css" type="text/css">
</head>

<body bgcolor="#DEF3FE" text="#000000">
<p><span class="detaillabels">修改密码完成，请重新登陆！</span></p>
<p> <span class="detaillabels"> 按<a href="chgpwd.asp?action=refresh">这里</a>或等待2秒钟</span><span class="detaillabels">返回。。。</span>
</body>
</html>
<%
else
%>
<html>
<head>
<title>更改密码</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="master.css" type="text/css">
</head>

<body bgcolor="#DEF3FE" text="#000000">
<div align="center">
  <p>&nbsp;</p>
<form action="chgpwd.asp" method="post">
    <table width="194" border="1" cellpadding="0" cellspacing="0">
      <tr> 
        <td class="detaillabels" width="92"> 
          <div align="center">登陆名</div>
        </td>
        <td class="detaillabels" width="96"> 
          <div align="center"><%=session("user_name")%></div>
        </td>
      </tr>
      <tr> 
        <td class="detaillabels" width="92"> 
          <div align="center">更改密码</div>
        </td>
        <td class="detaillabels" width="96"> 
          <div align="center"> 
            <input type="password" name="c_pwd" size="14">
          </div>
        </td>
      </tr>
      <tr> 
        <td class="detaillabels" colspan="2"> 
          <div align="center"> 
            <input type="submit" name="Submit" value="提交">
            <input type="reset" name="ReSet" value="重设">
          </div>
        </td>
      </tr>
    </table>
</form>
  <p>&nbsp;</p>
</div>
</body>
</html>
<%
	end if
end if
%>
