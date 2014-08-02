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
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
lid=request.QueryString("id")
RESPONSE.WRITE(UID)
RESPONSE.WRITE(LID)


if lid="" or  request.Cookies("adminuser")="" then
  response.write "来源未知或数据丢失!"
  response.end
End if

SXTS=request.form("SXTS")
if not(IsNumeric(SXTS)) THEN
     response.write "<script language=JavaScript>{window.alert('实休天数必须为正数 ！');window.location.href='XJ.asp';}</script>"
  RESPONSE.END
END IF
 
sql="select * from ZG where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    sqr1=rs1("zg")
end if
rs1.close


sql="select * from QXJDJB where id="&LID
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then 
rs("SXTS")=SXTS
rs("xJrq")=date()
rs("ZGZT")="2"
RS("XJR")=SQR1
   rs.update
   response.redirect("XJ.asp")
   response.end
end if
 
RESPONSE.WRITE(LID)



%>
