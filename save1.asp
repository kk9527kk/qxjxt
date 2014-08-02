<!--#include file="conn.asp"-->
<%
lid=request.QueryString("id")
dw=request.form("dw")
'xm=request.form("xm")
zcmc=request.form("zcmc")
xh=request.form("xh")
xlh=request.form("xlh")
pzqk=request.form("pzqk")
je=request.form("je")
zg=request.form("zg")

select case dw
case "局长室"
dwbm="10"
case "办公室"
dwbm="11"
case "法规科"
dwbm="12"
case "税政科"
dwbm="13"
case "征管科"
dwbm="14"
case  "计财科"
dwbm="15"
case "人教科"
dwbm="16"
case "监察室"
dwbm="17"
case "信息中心"
dwbm="18"
case "办税厅"
dwbm="19"
case "稽查局"
dwbm="20"
case  "永安一所"
dwbm="21"
case  "永安二所"
dwbm="22"
case  "草堂所"
dwbm="23"
case  "平皋所"
dwbm="24"
case  "竹园所"
dwbm="25"
case  "新民所"
dwbm="26"
case  "兴隆所"
dwbm="27"
case  "吐祥所"
dwbm="28"
case  "所得税科"
dwbm="29"

end select


if xh="" then
  xh="不详"
end if
if xlh="" then
  xlh="不详"
end if
if pzqk="" then
  pzqk="不详"
end if



if zg<>"" then
sql="select * from KP where id="&lid
response.write(sql)
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then
rs("dw")=dw
rs("dwbm")=dwbm
rs("zg")=zg
rs("zcmc")=zcmc
rs("xh")=xh
rs("xlh")=xlh
rs("pzqk")=pzqk
rs("je")=je
rs("djrq")=date()
rs("sl")=1
rs.update
response.redirect "index.asp"
end if
end if
%>
