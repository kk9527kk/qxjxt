<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%function isInteger(para) 
on error resume next 
dim str 
dim l,i 
if isNUll(para) then 
isInteger=false 
exit function 
end if 
str=cstr(para) 
if trim(str)="" then 
isInteger=false 
exit function 
end if 
l=len(str) 
for i=1 to l 
if mid(str,i,1)>"9" or mid(str,i,1)<"0" then 
isInteger=false 
exit function 
end if 
next 
isInteger=true 
if err.number<>0 then err.clear 
end function %>

<%
dim uid ,upwd
uid=trim(Request.Form("zh"))
zw=trim(Request.Form("zw"))
upwd=trim(Request.Form("mm"))
uid = Replace(uid, "'", "")
upwd = Replace(upwd, "'", "")
if not(isInteger(uid) and isInteger(upwd)) then
 	response.write "<script language=JavaScript>{window.alert('帐号和密码必须为数字!');window.location.href='login.asp';}</script>"
    response.end
end if 
if uid="" or upwd="" then
 	response.write "<script language=JavaScript>{window.alert('帐号和密码不能为空!');window.location.href='login.asp';}</script>"
    response.end
end if

set rs=server.createobject("adodb.recordset")
sqltext="select * from zg where id='" & uid & "' and mm='" & upwd & "'"
rs.open sqltext,conn,1,1
if rs.recordcount >= 1 then 
    
    response.cookies("flag") = "loginok"
    response.cookies("adminuser") = uid
    uid = rs("id")
    if rs("zw")="4" then
       Response.Redirect "index3.asp"
    elseif rs("zw")="2" then
       Response.Redirect "INDEX1.asp"
    elseif rs("zw")="3"  then
       Response.Redirect "INDEX2.asp"
    elseif rs("zw")="5" then
      Response.Redirect "index4.asp"
    elseif rs("zw")="6" then
      Response.Redirect "index5.asp"
    else
       Response.Redirect "index.asp"
    end if
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
