<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
'lid=request.QueryString("id")
uid=request.Cookies("adminuser")
'response.write(uid)
sql="select * from zg  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw1=rs1("dw")
    zg1=rs1("zg")
    zw1=rs1("zw")
    
end if
rs1.close
    response.cookies("flag") = "loginok"
    response.cookies("adminuser") = uid
    rqq=request.Cookies("rqq")
    rqz=request.Cookies("rqz")

%>

<html>
<%
 response.cookies("rqq") = rqq
 response.cookies("rqz") = rqz

rqq1=FormatDateTime(rqq,1)
rqz1=FormatDateTime(rqz,1)

%>


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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局请销假管理系统</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>



</head>

<body topmargin="2" background="IMAGES/header_1.gif" leftmargin="0">
<table border="0" id="table1" align="center"  cellspacing="0" cellpadding="0" width="759">
	<tr>
		<td width="757">
          <p align="center"><font color="#0000FF" face="黑体" size="6">
          云阳县国家税务局职工请销假查询清单</font>
        </td>
	</tr>
	<tr>
		<td width="757">
<p align="center"><b>申请时间:从<%=rqq1%>到<%=rqz1%>        </b>        </p>
        </td>
	</tr>
	<tr>
		<td width="757">
<br>
<p>　
        </td>
	</tr>
</table>

    <table border="0" width="775"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
         <td width="49" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        <font color="#008080">
        序号</font></td>
        <td width="77" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">申请人</td>
        <td width="70" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">休假时间起</td>
        <td width="106" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">休假时间止</td>
        <td width="90" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">请假申请日期</td>

		<td width="194" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center"> 审批状态</td>

		<td width="88" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center"> 销假日期</td>

             </tr>
            <% 
         bds="dw='"&dw1&"'"& "and sqsj between #"&rqq&"#"&" and #"&rqz&"#"  
             
          sql="select *  from qxjdjb  where "&bds &" order by zg desc"
          set rs=server.createobject("adodb.recordset")
          rs.open sql,connstr,1,1
       if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="7" width="769"> 
          <p align="center"><font color="#FF0000">暂时没有任何资料！</font></p>
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
         <td width="49" bgcolor="#FFFFFF" height="23" align="center"> 
         
          <% 
          qjzt=rs("qjzt")
          
        SELECT CASE QJZT
           CASE "1"
             QJZT1="已申请待科所长审批"
           CASE "2"
             QJZT1="科所长已审待人教科复核"
           CASE "3"
             QJZT1="人教科已审待分管局长审批"
           CASE "4"
             QJZT1="分管局长已审待局长终审"
           CASE "5"
             QJZT1="局长终审同意"
            CASE "6"
             QJZT1="人教科同意终审"
           CASE "7"
             QJZT1="分管局长同意终审"
           CASE "8"
             QJZT1="人教科同意待局长终审"
           CASE "9"
             QJZT1="不同意"
         END SELECT 
                     %>　
 
         
         
         
         
       <a href="cx.asp?id=<%=rs("id")%>"> <font color="#008080">
           <%=rs("id")%></font></a></td>
        <td width="77" bgcolor="#FFFFFF" height="23" align="center">    <%=rs("zg")%>　</td>
        <td width="70" bgcolor="#FFFFFF" height="23" align="center">　  <%=rs("qjsjq")%></td>
        <td width="106" bgcolor="#FFFFFF" height="23" align="center">   <%=rs("qjsjz")%>　</td>
        <td width="90" bgcolor="#FFFFFF" height="23" align="center">　  <%=rs("sqsj")%></td>
		<td width="194" bgcolor="#FFFFFF" height="23" align="center">　<%=QJZT1%> </td>
		<td width="88" bgcolor="#FFFFFF" height="23" align="center">　  <%=rs("XJRQ")%> </td>

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
      [<a href="dwtycxcl.asp?pageid=1">首页</a>] [<a href="dwtycxcl.asp?pageid=<%=k-1%>">上一页</a>]                                                                                                                   
      <%else%>
      [首页]&nbsp;[上一页] <%end if%>                                                                                                                  
      <%if k<>n then%>                                                                                                                  
      [<a href="dwtycxcl.asp?pageid=<%=k+1%>">下一页</a>] [<a href="dwtycxcl.asp?pageid=<%=n%>">尾页</a>]                                                                                                                   
      <%else%>
      [下一页]&nbsp;[尾页] <%end if%>                                                                                                                  
      共有<font color="red"><%=totalput%>                                                                                                                  
      </font>条记录 <b><font size="2" color="#0000FF">&nbsp;</font></b></td>                                                                                                                  
    <td height="23" width="302">
     
      <p align="right"><select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">
        <%                                                                                                                                                           
          i=1
          do while not i > n
          %>
        <option value="dwtycxcl.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>第 
        <%=i%> 页</option>
        <%
          i=i+1
          loop
        rs.close
        set rs=nothing

  %>
      </select></td>                                                                                                                
  </tr>
</table>


<center>
<p></p>
</center>

</body>       
       
</script>