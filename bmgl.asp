<!--#include file="conn.asp"-->
<html>
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
lid=request.QueryString("id")
'response.write(lid)
if  request.Cookies("adminuser")="" then
   response.write "来源未知或数据丢失!"
   response.end
end if
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
<title>云阳县国家税务局部门管理</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>
</head>

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

<table border="0" id="table1" align="center" cellpadding="0" width="713">
	<tr>
		<td width="711">
          <p align="center"><font color="#0000FF" face="黑体" size="6">
          云阳县国家税务局部门维护</font><p align="center">　
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="715" align="center" cellpadding="0" bgcolor="#000000" cellspacing="1">
      <tr> 
        <td width="131" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center">单位序号 
        </td>
        <td width="153" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        <p align="center">
        单位名称</p>
        </td>
        <td width="143" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">分管领导</td>
        <td width="143" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">管理</td>
              </tr>
      <%dim mc
    sql="select *  from dw order by dwbm  "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="4" width="352"> 
          <p align="center"><font color="#FF0000">暂时没有任何资料记录！</font></p>
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
        <td width="131" bgcolor="#FFFFFF" height="20" align="center"><%=rs("dwbm")%></td>
        <td width="153" bgcolor="#FFFFFF" height="20" align="center"><%=rs("dw")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="143"><%=rs("fgld")%></td>
        <td width="143" bgcolor="#FFFFFF" height="20" align="center"><a href="bmedit.asp?id=<%=rs("id")%>"><font color="#FF0000">修改</font></a></td>
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
  <table border="0" width="713" align="center" cellpadding="0">
    <tr>
    <td width="464" align="left">页数：<%=currentpage%>/<% =n%>
   	<%k=currentpage                                                   
   	if k<>1 then%>
   	[<a href="bmgl.asp?pageid=1">首页</a>]                                                   
   	[<a href="bmgl.asp?pageid=<%=k-1%>">上一页</a>]                                                   
   	<%else%>
   	[首页]&nbsp;[上一页]                                                   
   	<%end if%>
   	<%if k<>n then%>                                                   
   	[<a href="bmgl.asp?pageid=<%=k+1%>">下一页</a>]                                                   
   	[<a href="bmgl.asp?pageid=<%=n%>">尾页</a>]                                                   
   	<%else%>
   	[下一页]&nbsp;[尾页]                                                   
   	<%end if%>
        共有<font color="red"><%=totalput%></font>条记录                                                  
      <b><font color="#0000FF"><font size="2">&nbsp;</font></font></b></td> 
          <form action="bmgl.asp" name="pfrm">     
      <td height="23" width="245" align="left">  
        <p align="right">
        <select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">  
          <%                              
          i=1
          do while not i > n
          %>
          <option value="bmgl.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>第 
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
       <tr>
      <td width="711" colspan="2" align="left">                        
      </td>                     
    </tr>         
          
    <tr>         
      <td width="711" height="35" colspan="2" valign="bottom" align="left"> 
        　</td>        
    </tr>        
  </table>        
</div>        
  <p>
        　
  </p>
</body>       
       
</html>       
</script>