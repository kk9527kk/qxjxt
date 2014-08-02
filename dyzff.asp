<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='index.htm';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
'response.write(uid)

sql="select * from djb  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
    zw=rs1("zw")
end if
rs1.close
%>

<html>

<head>
<STYLE type=text/css>A:link {
	COLOR: #000000; FONT-FAMILY: 宋体; TEXT-DECORATION: none
}
A:visited {
	COLOR: #000000; FONT-FAMILY: 宋体; TEXT-DECORATION: none
}
A:active {
	FONT-FAMILY: 宋体; TEXT-DECORATION: underline
}
A:hover {
	COLOR: #84bd6b; TEXT-DECORATION: underline overline
}
BODY {
	COLOR: #000000; FONT-FAMILY: 宋体; FONT-SIZE: 9pt
}
TABLE {
	COLOR: #000000; FONT-FAMILY: 宋体; FONT-SIZE: 9pt
}
.f24 {
	COLOR: #ff0000; FONT-FAMILY: 宋体; FONT-SIZE: 9pt; TEXT-DECORATION: underline overline
}
.l17 {
	LINE-HEIGHT: 170%
}
.f18 {
	FONT-SIZE: 18px
}
</STYLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局计算机配件及耗材审批管理</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>
</head>

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="880" height="47">
	<tr>
		<td width="878" height="35">
          <p align="center"><font color="#0000FF" face="黑体" size="6">云阳县国家税务局打印纸发放</font>
          <p align="center">　
        </td>
	</tr>
	<tr>
		<td width="878" height="12">
        </td>
	</tr>
</table>
<div align="left">
        <table border="0" width="764"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000" height="64">
      <tr> 
        <td width="100" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center"> 
          <p align="center">库存单位 
        </td>
        <td width="71" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">
        打印纸型号</td>
        <td width="94" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">库存数量</td>
        <td width="121" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">发放单位</td>
        <td width="216" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">
        发放数量(包)</td>
		<td width="109" bgcolor="#FFFFFF" height="21" background="images/bg1.gif" align="center">
        发放处理</td>

      
  <%dim mc
    sql="select *  from SQSPB WHERE ( qs='1' and  hcmc='打印纸' and kcbz='1' and qssl>0) order by SQRQ,id "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="6" width="438" height="14"> 
          <p align="center"><font color="#FF0000">暂时没有任何待发放的资料！</font></p>
        </td>
      </tr>
      <%
else
const maxperpage=10 '定义每一页显示的数据记录的常量
dim currentpage '定义当前页的变量
rs.pagesize=maxperpage
currentpage=request.querystring("pageid")
if currentpage="" then
currentpage=1
elseif currentpage<1 then
currentpage=1
else
currentpage=clng(currentpage)
	if currentpage > rs.pagecount then
	currentpage=rs.pagecount
	end if
end if
'如果变量currentpage的数据类型不是数值型
'就1赋给变量currentpage
if not isnumeric(currentpage) then
currentpage=1
end if
dim totalput,n '定义变量
totalput=rs.recordcount
if totalput mod maxperpage=0 then
n=totalput\maxperpage
else
n=totalput\maxperpage+1
end if
if n=0 then
n=1
end if
rs.move(currentpage-1)*maxperpage
i=0
do while i< maxperpage and not rs.eof
%>
      <tr> 
        <td width="100" bgcolor="#FFFFFF" height="23" align="center"><%=rs("dw")%></td>
        <td width="71" bgcolor="#FFFFFF" height="23" align="center"><%=rs("xh")%></td>
        <td bgcolor="#FFFFFF" height="23" align="center" width="94"><%=rs("qssl")%>　</td>
        <td bgcolor="#FFFFFF" height="23" align="center" width="121">
         <form method="POST" action="dyzffcl.asp?id=<%=rs("id")%>">
            <select size="1" name="ffdw">
              
              
              <option value="局长室">局长室</option>
            <option value="办公室">办公室</option>
            <option value="法规科">法规科</option>
            <option value="税政科">税政科</option>
            <option value="税政科">税政科</option>
            <option value="所得税科">所得税科</option>
            <option value="征管科">征管科</option>
            <option value="计财科">计财科</option>
            <option value="人教科">人教科</option>
            <option value="监察室">监察室</option>
            <option value="信息中心">信息中心</option>
            <option value="办税厅">办税厅</option>
            <option value="稽查局">稽查局</option>
            <option value="永安一所">永安一所</option>
            <option value="永安二所">永安二所</option>
            <option value="草堂所">草堂所</option>
            <option value="平皋所">平皋所</option>
            <option value="竹园所">竹园所</option>
            <option value="新民所">新民所</option>
            <option value="兴隆所">兴隆所</option>
            <option value="吐祥所">吐祥所</option>
              
              
              </select>
            　
           

</td>
        <td bgcolor="#FFFFFF" height="23" align="center" width="216"><input type="text" name="ffSL" size="12">
        </td>
        <td bgcolor="#FFFFFF" height="23" align="center" width="109"><input type="submit" value="发放" name="B1"></td>   </form>
       
        </form> 
      </tr>
       <%
      i=i+1
    rs.movenext
    loop
    end if
    %>

    </table>
<table border="0" width="764" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td width="464">页数：<%=currentpage%>
      /<% =n%>                                                                                            
      <%k=currentpage                                                                                                                                                           
   	if k<>1 then%>
      [<a href="sq.asp?pageid=1">首页</a>] [<a href="sq.asp?pageid=<%=k-1%>">上一页</a>]                                                                                             
      <%else%>
      [首页]&nbsp;[上一页] <%end if%>                                                                                            
      <%if k<>n then%>                                                                                            
      [<a href="sq.asp?pageid=<%=k+1%>">下一页</a>] [<a href="sq.asp?pageid=<%=n%>">尾页</a>]                                                                                             
      <%else%>
      [下一页]&nbsp;[尾页] <%end if%>                                                                                            
      共有<font color="red"><%=totalput%>                                                                                            
      </font>条记录 <b><font size="2" color="#0000FF">&nbsp;</font></b></td>                                                                                            
    <td height="23" width="302">
    <form action="sq.asp" name="pfrm">   
      <p align="right"><select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">
        <%                                                                                                                                       
          i=1
          do while not i > n
          %>
        <option value="sq.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>第 
        <%=i%> 页</option>
        <%
          i=i+1
          loop
          %>
      </select></td>                                                                                            
  </tr></form>
</table>

</html>



