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
'if not(isInteger(uid) and isInteger(upwd)) then
'if not isInteger(uid) then

' 	response.write "<script language=JavaScript>{window.alert('ÕÊºÅºÍÃÜÂë±ØĞëÎªÊı×Ö!');window.location.href='login.asp';}</script>"
'    response.end
'end if 
if uid="" or upwd="" then
 	response.write "<script language=JavaScript>{window.alert('ÕÊºÅºÍÃÜÂë²»ÄÜÎª¿Õ!');window.location.href='login.asp';}</script>"
    response.end
end if

set rs=server.createobject("adodb.recordset")
sqltext="select * from zg where id='" & uid & "' and mm='" & upwd & "'"
rs.open sqltext,conn,1,1
if rs.recordcount >= 1 then 
    response.cookies("flag") = "loginok"
    response.cookies("adminuser") = uid
    uid = rs("id")
    select case rs("zw")
      case "1" 
       Response.Redirect "index1.asp"
      case "2"
         Response.Redirect "index1.asp"
      case "3"
          Response.Redirect "index3.asp"
      case "4"
          Response.Redirect "index4.asp"
      case "5"
          Response.Redirect "index5.asp"
      case "6"
          Response.Redirect "index6.asp"
      case "7"
          Response.Redirect "index7.asp"
    end select
     
      rs.close
else
 	response.write "<script language=JavaScript>{window.alert('ÕÊºÅºÍÃÜÂë´íÎó!');window.location.href='login.asp';}</script>"
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
