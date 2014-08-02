<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>

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
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='index.htm';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
sql="select * from djb  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw=rs1("dw")
    zg=rs1("zg")
end if
rs1.close

%>


<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局计算机耗材及配件申请审批表</title>
</head>

<body background="images/header_1.gif" bgproperties="fixed">

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="764">
	<tr>
		<td width="762">
          <p align="center"><font color="#0000FF" face="黑体" size="6">云阳县国家税务局打印纸签收</font>
          <p align="center">　
        </td>
	</tr>
</table>
    <table border="0" width="774"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="94" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center">申请时间 
        </td>
        <td width="102" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">签收人</td>
        <td width="123" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">耗材名称</td>
		<td width="120" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        申请数量</td>

        <td width="136" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        是否终审</td>

       
        <td width="119" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        使用人</td>

       
        <td width="195" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        型号</td>

       
        <td width="98" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">签收数量</td>
       
       
        <td width="94" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">签收金额</td>
       
       
        <td width="113" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">处理</td>
       
      
      
  <%dim mc
    sql="select *  from SQSPB WHERE dw='"&dw&"' order by id desc "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
      <tr> 
        <td bgcolor="#FFFFFF" colspan="10" width="438"> 
          <p align="center"><font color="#FF0000">暂时没有任何待签收的资料！</font></p>
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
        <td width="94" bgcolor="#FFFFFF" height="20" align="center"><%=rs("sqrq")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="102"><%=rs("qsr")%>　</td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="123"><%=rs("hcmc")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="120"><%=rs("sl")%></td>
        <td width="136" bgcolor="#FFFFFF" height="20" align="center">
        
        <%if rs("zt")="4" then %>
        <font color="#FF0000">终审并同意</font>  
        <%elseif rs("zt")="5" or rs("zt")="6" or rs("zt")="7"  then%>   
        <font color="#008000">终审不同意 </font>    
        <%else%>
         <font color="#008880">未终审 </font>    
        <%end if%>
        </td>
        <td width="119" bgcolor="#FFFFFF" height="20" align="center"><form method="POST" action="qs.asp?id=<%=rs("id")%>" NAME="FORM1">
        <%if (rs("zt")="4" and trim(rs("qs"))<>"1") then %>
       
<input type="text" name="syr" size="7" value="<%=RS("syr")%>">
         <%else%>                      
            <%=RS("syr")%>                    
         <%END IF%>           

        
        
        </td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="195"><%if (rs("zt")="4" and trim(rs("qs"))<>"1") then %>
       
<input type="text" name="xh" size="12" value="<%=RS("xh")%>">
         <%else%>                
            <%=RS("xh")%>              
         <%END IF%>     
        </td>
        <td width="98" bgcolor="#FFFFFF" height="20" align="center">
         <%if (rs("zt")="4" and trim(rs("qs"))<>"1") then %>
               
            <input type="text" name="QSSL" size="8" value="<%=RS("qsSL")%>">
          <%else%>              
            <%=RS("qsSL")%>            
         <%END IF%>                                            　</td>
  
        <td width="94" bgcolor="#FFFFFF" height="20" align="center">
            
             <%if (rs("zt")="4" and trim(rs("qs"))<>"1") then %>
        <input type="text" name="QSJE" size="9" value="<%=RS("JE")%>">
        <%else%>     
            <%=RS("je")%>   

         </td>
       <%END IF%>
        <td width="113" bgcolor="#FFFFFF" height="20" align="center" valign="middle">
        
        <%if (rs("zt")="4" and trim(rs("qs"))<>"1") then %>
        <input type="submit" value="签收" name="B1">   
         <%elseif rs("zt")="4" and rs("qs")="1"  then%>                              
         <font color="#008880">已收 </font>                                 
         <%else%>
         <font color="#008880"></font>                                 
        <%end if%>
              
        
        </td>
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
      [<a href="qs1.asp?pageid=1">首页</a>] [<a href="qs1.asp?pageid=<%=k-1%>">上一页</a>]                                                                                                  
      <%else%>
      [首页]&nbsp;[上一页] <%end if%>                                                                                                 
      <%if k<>n then%>                                                                                                 
      [<a href="qs1.asp?pageid=<%=k+1%>">下一页</a>] [<a href="qs1.asp?pageid=<%=n%>">尾页</a>]                                                                                                  
      <%else%>
      [下一页]&nbsp;[尾页] <%end if%>                                                                                                 
      共有<font color="red"><%=totalput%>                                                                                                 
      </font>条记录 <b><font size="2" color="#0000FF">&nbsp;</font></b></td>                                                                                                 
    <td height="23" width="302">
    <form action="qs1.asp" name="pfrm">   
      <p align="right"><select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 16; width: 77; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">
        <%                                                                                                                                            
          i=1
          do while not i > n
          %>
        <option value="qs1.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>第 
        <%=i%> 页</option>
        <%
          i=i+1
          loop
          %>
      </select></td>                                                                                                 
  </tr></form>
</table>

</html>
