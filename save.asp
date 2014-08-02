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
lid=request.QueryString("id")
if lid="" or lid<>request.Cookies("adminuser") then
response.write "来源未知或数据丢失!"
response.end
end if
qjsjq=request.form("qjsjq")
qjsjz=request.form("qjsjz")
ts=request.form("ts")
qjlb=request.form("qjlb")
qjsx=request.form("qjsx")
qjlb=request.form("qjlb")
QWHD=REQUEST.FORM("QWHD")
SFLFJ=REQUEST.FORM("SFLFJ")
IF SFLFJ="YES" THEN
  SFLFJ="YES"
 ELSE
  SFLFJ="NO"
 END IF
 if DateDiff("d",qjsjq,qjsjz)<0 then 
   response.write "<script language=JavaScript>{window.alert('请假止时间应晚于请假起时间，请重输入！');window.location.href='sq.asp';}</script>"
   response.end
end if

if qjsjq="" then
  response.write "<script language=JavaScript>{window.alert('请假时间起不允许为空！');window.location.href='SQ.asp';}</script>"
response.end
end if
if qjsjz="" then
  response.write "<script language=JavaScript>{window.alert('请假时间止不允许为空！');window.location.href='SQ.asp';}</script>"
response.end
end if

if qjsx="" then
  response.write "<script language=JavaScript>{window.alert('请假事项不允许为空！');window.location.href='SQ.asp';}</script>"
response.end
end if

if qjsX="" then
  response.write "<script language=JavaScript>{window.alert('请假事项为空！');window.location.href='SQ.asp';}</script>"
response.end
end if

if (not isInteger(TS)) then
   response.write "<script language=JavaScript>{window.alert('天数必须为自然数！');window.location.href='SQ.asp';}</script>"
response.end
end if

if QWHD="" then
  response.write "<script language=JavaScript>{window.alert('前往何地不允许为空！');window.location.href='SQ.asp';}</script>"
response.end
end if


'sqr=request.cookies("adminuser")

sql="select * from zg  where id='"&lid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg1=rs1("zg")
    zw=rs1("zw")
 '   fgld=rs1("fgld")
end if
sxts=0
select case zw
  case "1"
    qjzt="1"
    message="申请已经成功提交,等待科所领导审批"
  case  "2"
    qjzt="1"
     message="申请已经成功提交,等待科所领导审批"
  case "3"
    qjzt="3"
    message="申请已经成功提交,等待局领导审批"
  case  "4"
    qjzt="4"
    message="申请已经成功提交,等待局长审批"
  case "5"
    qjzt="5"
    sxts=ts
    message="申请已经成功保存，并已终审！"
  case  "6"
    qjzt="2"
    message="申请已经成功提交,等待人教科审批"
  case "7" 
    qjzt="1"
    message="申请已经成功提交,等待科所长审批"
 end select
rs1.close
sql="select * from qxjdjb"
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
rs.addnew
rs("qjsx")=qjsx
rs("qjsjq")=qjsjq
rs("qjsjz")=qjsjz
rs("qjlb")=qjlb
rs("ts")=ts
rs("qwhd")=qwhd
rs("sqsj")=date()
rs("sxts")=sxts
rs("dw")=dw
rs("zg")=zg1
rs("qjzt")=qjzt
rs("zgzt")="2"
rs("SFLFJ")=SfLFJ
rs("QJRZW")=zw
rs.update
response.write "<script language=JavaScript>{window.alert('"&message&"');window.location.href='SQ.asp';}</script>"
    response.end
%>
