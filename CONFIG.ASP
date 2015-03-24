<%
    path="database/"     '目录
    db="northwind.mdb"    '数据库文件名
    strpath=server.mappath(path&db)    '转换绝对路径
    constr="Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & strPath
%>
