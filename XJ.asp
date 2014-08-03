

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
<title>云阳县国家税务局请销假管理</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>
</head>

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

&nbsp;           

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->            

<%            
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
sql="select * from ZG  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
    zw=rs1("ZW")
end if
rs1.close
%>

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="753">
	<tr>
		<td width="751">
          <p align="center"><font color="#0000FF" face="黑体" size="6">
          云阳县国家税务局销假处理</font>
        </td>
	</tr>
	<tr>
		<td width="751">
          　
          <p>　
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="761"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="85" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center"><a href="index2.asp">单位</a> 
        </td>
        <td width="93" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        请请人</td>
        <td width="99" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">请假时间起</td>
        <td width="150" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        请假时间止</td>
		<td width="120" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        请假天数</td>

        <td width="118" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        申请时间</td>

       
        <td width="78" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">实休天数</td>
       
       
        <td width="95" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">处理</td>
       
      </tr>
      <%dim mc
    	select case ZW
		case "3"
			sql="select *  from QXJDJB WHERE (QJZT='5' OR QJZT='6' OR QJZT='7') AND ZGZT='1' and dw='"&dw&"' order by SQSJ "
		case "6"
			sql="select *  from QXJDJB WHERE (QJZT='5' OR QJZT='6' OR QJZT='7') AND ZGZT='1' and zw=3 order by SQSJ "
		case "7"
			sql="select *  from QXJDJB WHERE (QJZT='5' OR QJZT='6' OR QJZT='7') AND ZGZT='1' and zw=3 order by SQSJ "
		case "5"
			sql="select *  from QXJDJB WHERE (QJZT='5' OR QJZT='6' OR QJZT='7') AND ZGZT='1' and zw=4 order by SQSJ "
		end select
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="8" width="719"> 
          <p align="center"><font color="#FF0000">暂时没有任何待销假的资料！</font></p>
        </td>
      </tr>
      <%
else
const maxperpage=25 '定义每一页显示的数据记录的常量
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
        <td width="85" bgcolor="#FFFFFF" height="20" align="center"><%=rs("dw")%></a></td>
        <td width="93" bgcolor="#FFFFFF" height="20" align="center"><%=rs("ZG")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="99"><%=rs("QJSJQ")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="150"><%=rs("QJSJZ")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="120"><%=rs("TS")%></td>
        <td width="118" bgcolor="#FFFFFF" height="20" align="center"><a href="cx.asp?id=<%=rs("id")%>"><font color="#FF0000">&nbsp;&nbsp;<%=rs("SQSJ")%></font></a></td>
      
        <td width="78" bgcolor="#FFFFFF" height="20" align="center">
       <form method="POST" action="XJCL.asp?id=<%=rs("id")%>" name="form1">
           <input type="text" name="SXTS" size="7" value="<%=rs("TS")%>">　
          </td>
       
        <td width="95" bgcolor="#FFFFFF" height="20" align="center" valign="middle"><font color="#FF0000"><input type="submit" value="销假" name="B1"></font></form></td>
       
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
  <table border="0" width="762" align="center"  cellspacing="0" cellpadding="0">
    <tr>
    <td width="464">页数：<%=currentpage%>/<% =n%>
   	<%k=currentpage                                                                                
   	if k<>1 then%>
   	[<a href="XJ.asp?pageid=1">首页</a>]                                                                                
   	[<a href="XJ.asp?pageid=<%=k-1%>">上一页</a>]                                                                                
   	<%else%>
   	[首页]&nbsp;[上一页]                                                                                
   	<%end if%>
   	<%if k<>n then%>                                                                                
   	[<a href="XJ.asp?pageid=<%=k+1%>">下一页</a>]                                                                                
   	[<a href="XJ.asp?pageid=<%=n%>">尾页</a>]                                                                                
   	<%else%>
   	[下一页]&nbsp;[尾页]                                                                                
   	<%end if%>
        共有<font color="red"><%=totalput%></font>条记录                                                                               
      <b><font color="#0000FF"><font size="2">&nbsp;</font></font></b></td> 
            
      <td height="23" width="294">  
        <p align="right">
        <select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">  
          <%                                                            
          i=1
          do while not i > n
          %>
          <option value="XJ.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>第 
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
      <td width="760" colspan="2">             
      </td>                     
    </tr>         
    <%end if%>         
    <tr>         
      <td width="760" height="35" colspan="2" valign="bottom"> 
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