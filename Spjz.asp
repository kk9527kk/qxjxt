<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='index.htm';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
'response.write(uid)

sql="select * from ZG  where id='"&uid&"'"
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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局计算机配件及耗材审批管理</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>
</head>

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="769">
	<tr>
		<td width="767">
          <p align="center"><font color="#0000FF" face="黑体" size="6">
          云阳县国家税务局请销假审批事项</font>
          <p align="center">　
        </td>
	</tr>
	<tr>
		<td width="767">
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="772"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="50" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center">单位 
        </td>
        <td width="60" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        请假人</td>
        <td width="60" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        请假时间起</td>
        <td width="68" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">请假时间止</td>
        <td width="49" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        请假天数</td>
		<td width="54" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        申请时间</td>

        <td width="147" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        <font size="2">请假状态</a></font></td>

       
        <td width="60" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">审批处理</td>
       
      </tr>
      <%dim mc
      
      
   sql="select *  from QXJDJB WHERE (QJZT='4' or qjzt='8')  order by SQSJ"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="8" width="446"> 
          <p align="center"><font color="#FF0000">暂时没有任何待审批的资料！</font></p>
        </td>
      </tr>
      <%
else
const maxperpage=15 '定义每一页显示的数据记录的常量
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
        <td width="50" bgcolor="#FFFFFF" height="20" align="center"> <%=rs("dw")%></td>
        <td width="60" bgcolor="#FFFFFF" height="20" align="center"><%=rs("ZG")%></td>
        <td width="60" bgcolor="#FFFFFF" height="20" align="center"><%=rs("QJSJQ")%>　</td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="68"><%=rs("QJSJZ")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="49"><%=rs("TS")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="54"><%=rs("SQSJ")%></td>
        <%
        QJZT=RS("QJZT")
        SELECT CASE QJZT
           CASE "1"
             QJZT1="已申请待科所长审批"
           CASE "2"
             QJZT1="科所长已审待人教科复核"
           CASE "3"
             QJZT1="人教科已审待分管局长审"
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
        
        <td width="147" bgcolor="#FFFFFF" height="20" align="center">　<%=QJZT1%></td>
      
        <td width="60" bgcolor="#FFFFFF" height="20" align="center"><a href="spjzk.asp?id=<%=rs("id")%>"><font color="#FF0000">审批</font></a></td>
       
        <%
      i=i+1
    rs.movenext
    loop
    end if
    %>
      </tr>
    </table>
</div>
<div align="left">
  <table border="0" width="771" align="center"  cellspacing="0" cellpadding="0">
    <tr>
    <td width="464">页数：<%=currentpage%>/<% =n%>
   	<%k=currentpage                                                    
   	if k<>1 then%>
   	[<a href="spjz.asp?pageid=1">首页</a>]                                                    
   	[<a href="spjz.asp?pageid=<%=k-1%>">上一页</a>]                                                    
   	<%else%>
   	[首页]&nbsp;[上一页]                                                    
   	<%end if%>
   	<%if k<>n then%>                                                    
   	[<a href="spjz.asp?pageid=<%=k+1%>">下一页</a>]                                                    
   	[<a href="spjz.asp?pageid=<%=n%>">尾页</a>]                                                    
   	<%else%>
   	[下一页]&nbsp;[尾页]                                                    
   	<%end if%>
        共有<font color="red"><%=totalput%></font>条记录                                                   
      <b><font color="#0000FF"><font size="2">&nbsp;</font></font></b></td> 
          <form action="spjz.asp" name="pfrm">     
      <td height="23" width="303">  
        <p align="right">
        <select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">  
          <%                              
          i=1
          do while not i > n
          %>
          <option value="spjz.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>第 
          <%=i%> 页</option>
          <%
          i=i+1
          loop
          %>
        </select>                              

        </td>
    </tr>
    </form>
  <center>
    <%if session("admin")<>"" then%>
    <tr>
      <td width="769" colspan="2">             
      </td>                     
    </tr>         
    <%end if%>         
    <tr>         
      <td width="769" height="35" colspan="2" valign="bottom"> 
        　</td>        
    </tr>        
  </table>        
</div>        
  <p>
        　
  </p>
</center>
</body>       
       
</html>       
</script>