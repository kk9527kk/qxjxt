<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<html>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>zhanshi</title>

</head>

<table border="1">



<% 
hhh =  FormatDateTime(now, vbShortDate)      
sql="select *  from qxjdjb  where QJZT  >= '5' and QJSJQ <=  #"&hhh&"# and QJSJZ >= #"&hhh&"#  order by id desc"
set rs=server.createobject("adodb.recordset")
//response.write sql
rs.open sql,connstr,1,1
do while not rs.eof
%>



<tr>
<td><%=rs("ZG")%></td>
<td><%=rs("qjlb")%></td>
<td><%=rs("QJSJQ")%></td>
<td><%=rs("QJSJZ")%></td>
</tr>
<%
rs.movenext
loop
%>
</table>

</html>