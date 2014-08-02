<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
dim uid ,upwd
uid=trim(Request.Form("zh"))
upwd=trim(Request.Form("mm"))
uid = Replace(uid, "'", "")
upwd = Replace(upwd, "'", "")
if uid="" or upwd="" then
 	response.write "<script language=JavaScript>{window.alert('帐号和密码不能为空!');window.location.href='login.asp';}</script>"
    response.end
end if

set rs=server.createobject("adodb.recordset")
sqltext="select * from djb where id='" & uid & "' and mm='" & upwd & "'"
rs.open sqltext,conn,1,1
if rs.recordcount >= 1 then 
    if zw="1" then
      uid = rs("id")
    end if
    response.cookies("flag") = "loginok"
    response.cookies("adminuser") = uid
    Response.Redirect "sq.asp"
    rs.close
else
 	response.write "<script language=JavaScript>{window.alert('帐号和密码错误!');window.location.href='login.asp';}</script>"
    response.end
    rs.close
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body>
</body>
</html>
