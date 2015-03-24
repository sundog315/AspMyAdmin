<%if Session("admin")=true then%>
<!--#include file="config.asp"-->
<%
on error resume next
Sub RenameFile(sFrom, sTo)
	Dim fso, f
	Set fso = CreateObject("Scripting.FileSystemObject")
	sDateTimeStamp = "." & year(now) & month(now) & day(now) & "-" & hour(now) & minute(now) & second(now)
	fso.MoveFile sTo, Server.MapPath(path & db & sDateTimeStamp & ".bak")
	fso.MoveFile sFrom, sTo
End Sub

sConnFrom = constr
sConnTo = "Provider=Microsoft.Jet.OLEDB.4.0;" &_
	 "Data Source=" & Server.MapPath(path & "tempdb.mdb")

sFileFrom = Server.MapPath(path & db)
sFileTo = Server.MapPath(path & "tempdb.mdb")

if session("admin")=true then
	set jro = server.createobject("jro.JetEngine")
	jro.CompactDatabase sConnFrom, sConnTo
			
	if err <> 0 then
		bError = True
		response.write "Error:<br>" & err.description
	end if
	
	RenameFile sFileTo, sFileFrom
			
	if err <> 0 then
		bError = True
		response.write "Error:<br>" & err.description
	end if
			
	if not bError then
%>
<html>
<head>
<meta HTTP-EQUIV="Refresh" content="2; URL=main.asp">
<title>压缩数据库</title>
<link rel="stylesheet" href="master.css" type="text/css">
</head>

<body bgcolor="#DEF3FE">
<p class="detaillabels"><br>
  压缩成功！<br>
</p>
<p class="sundog>压缩完成！</p><br>
<span class="sundog"><span class="detaillabels">原数据库被更名为"数据库名.日期.bak" 
  <%
	else
%>
  <br>
  <br>
  压缩时产生错误！ <br>
  <%
	end if
else
   response.redirect "default.asp"
end if
%>
  </span>
<%end if%>
