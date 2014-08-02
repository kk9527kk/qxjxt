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
QSSL=REQUEST.FORM("QSSL")
QSJE=REQUEST.FORM("QSJE")
qsxh=REQUEST.FORM("xh")
qssyr=REQUEST.FORM("syr")

if qsxh="" then
QSxh="不详"
END IF

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

sql="select * from djb  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
   dw=rs1("dw")
   zg=rs1("zg")
   zw=rs1("zw")
end if

rs1.close
if qssyr="" then
QSsyr=zg
END IF

if not isNumeric(qsje)  then
response.write "<script language=JavaScript>{window.alert('签收金额必须为不小于0的实数!');window.location.href='qS1.asp';}</script>"
response.end
end if

if IsNumeric(qsje) THEN
       if qsje<0 then
     response.write "<script language=JavaScript>{window.alert('签收金额不能为负数！');window.location.href='qS1.asp';}</script>"
  RESPONSE.END
  end if
END IF

if not isInteger(qssl) then
response.write "<script language=JavaScript>{window.alert('签收数量必须为正整数!');window.location.href='QS1.asp';}</script>"
response.end
end if

if  QSSL<=0  then
    
response.write "<script language=JavaScript>{window.alert('签收数量必须为正整数!');window.location.href='QS1.asp';}</script>"
response.end
end if


sql="select * from sqspb where id="&lid
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then 
rs("qs")="1"
rs("xh")=qsxh
rs("syr")=qssyr
RS("qsr")=ZG
IF (RS("kcbz")<>"1" AND  RS("HCMC")="打印纸") THEN
 rs("jsrq")=date() 
 rs("sqr")=ZG
else
'rs("sqr")=ZG
 RS("QSRQ")=DATE() 
 RS("QSSL")=QSSL
 RS("JE")=QSJE
END IF 
    rs.update
   response.write "<script language=JavaScript>{window.alert('已签收成功!');window.location.href='qs1.asp';}</script>"
   response.end
 
end if
 


%>
