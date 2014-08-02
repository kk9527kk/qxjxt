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
if lid="" or  request.Cookies("adminuser")="" then
   response.write "来源未知或数据丢失!"
   response.end
end if

qssl=request.form("ffsl")
dw=request.form("ffdw")

if trim(qssl)="" then
   response.write "<script language=JavaScript>{window.alert('分发数量不能空 ！');window.location.href='dyzff.asp';}</script>"
RESPONSE.END
end if
    
  if not(isInteger(qssl) ) THEN
     response.write "<script language=JavaScript>{window.alert('分发数量应为自然数！');window.location.href='dyzff.asp';}</script>"
  RESPONSE.END
END IF
  if qssl<=0 then
   response.write "<script language=JavaScript>{window.alert('分发数量不能小于或等于0 ！');window.location.href='dyzff.asp';}</script>"
  RESPONSE.END
END IF  
   
   


sql="select * from sqspb where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then 

'response.write(qssl)
'response.write(rs("qssl"))
   if qssl-rs("qssl")>0 then
    response.write "<script language=JavaScript>{window.alert('分发数量不能大于库存数量 ！');window.location.href='dyzff.asp';}</script>"
    RESPONSE.END
   END IF 
je=rs("je")
qssl1=rs("qssl")
xh=rs("xh")
hcmc=rs("hcmc")
qsrq=rs("qsrq")
xhbz=rs("xhbz")
xhrq=rs("xhrq")
'response.write(je)
'response.write(qssl1)
rs("je")=je*(qssl1-qssl)/qssl1
rs("qssl")=rs("qssl")-qssl
'rs("xhrq")=date()
  'rs("xhbz")="1"
   rs.update
  ' response.redirect("dyzff.asp")
 '  response.end

rs.close
set rs=nothing
end if

sql="select * from sqspb"
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
rs.addnew
rs("hcmc")=hcmc
rs("xh")=xh
'rs("sqsx")=sqsx
rs("sl")=qssl
rs("QSsl")=qssl
rs("sqrq")=date()
rs("QSrq")=date()
rs("JE")=je*Qssl/qssl1
rs("dw")=dw
'rs("sqr")=sqr1
rs("zt")="4"
rs("xhbz")=xhbz
rs("xhrq")=xhrq
rs("qs")="0"
rs("kcbz")="0"
rs("qsrq")=qsrq
rs("jldrq")=date()
rs.update
response.write "<script language=JavaScript>{window.alert('已经成功分发！');window.location.href='dyzff.asp';}</script>"
    response.end

%>
