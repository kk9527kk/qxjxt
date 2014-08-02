<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file ="conn.asp"-->
<%
    



 dw1=request.QueryString("dw")
if dw1<>"" then
dw=mid(dw1,2,len(dw1)-2)
end if
aa=request.QueryString("hcmc")
rqq=request.Cookies("rqq")
rqz=request.Cookies("rqz")
rqq1=FormatDateTime(rqq,1)
rqz1=FormatDateTime(rqz,1)
if aa="" and dw1="" then
  fjgs="云阳县国家税务计算机耗材领用清单"
elseif aa="" and dw1<>"" then
  fjgs="云阳县国家税务"&dw&"计算机耗材领用清单"
elseif aa<>"" and dw1="" then
  aa1=mid(aa,2,len(aa)-2)
  fjgs="云阳县国家税务"&aa1&"领用清单"
elseif aa<>"" and dw1<>"" then
    aa1=mid(aa,2,len(aa)-2)
  fjgs="云阳县国家税务局"&dw&aa1&"领用清单"
end if
    response.cookies("rqq") = rqq
    response.cookies("rqz") = rqz
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
<title>云阳县国家税务局计算机资产管理</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>
</head>

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

<table border="0" id="table1"  align="center"  cellspacing="0" cellpadding="0" width="777">
	<tr>
		<td width="775">
          <p align="center"><b><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
          </font><font color="#000080" face="黑体" size="6"><%=fjgs%></font></b></p>
                    <p align="center"><font face="黑体" size="2"><b>统计时间：从<%=rqq1%>到<%=rqz1%></b></font></p>
          <p align="right"><font color="#000080" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;   
          <a href="cxquit.asp">返回上页</a> &nbsp;&nbsp;&nbsp;</b><b> </b></font></p>    
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="765"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="63" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center"> 
          <p align="center">单位 
        </td>
        <td width="53" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        使用人</td>
        <td width="69" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">耗材名称</td>
        <td width="116" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        型号</td>
		<td width="48" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        数量</td>

        <td width="71" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        <font size="2">申请日期</font></td>

        <td width="73" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        配置日期</td>

        <td width="47" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        金额</td>

        
      </tr>
  


      <%if (aa="" and dw1="") then 
         bds="xhbz='1'    and xhrq between #"&rqq&"#"&" and #"&rqz&"#"
       end if
       if (aa<>"" and dw1="" ) then           
         bds="hcmc="&aa&" and xhbz='1' and xhrq between #"&rqq&"#"&" and #"&rqz&"#"
       end if 
       if (aa="" and dw1<>"" ) then           
         bds="dw="&dw1&"  and xhbz='1' and xhrq between #"&rqq&"#"&" and #"&rqz&"#"
       end if 
       if (aa<>"" and dw1<>"") then
         bds="hcmc="&aa&" and dw="&dw1&"  and xhbz='1'  and xhrq between #"&rqq&"#"&" and #"&rqz&"#"
       end if

      sql="select *  from sqspb where "&bds&" order by id desc"   
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    if rs.eof then
      %>
                   
      <tr> 
        <td width="595" bgcolor="#FFFFFF" height="20" align="center" colspan="8"> 
          <p align="center"><font color="#FF0000">暂时没有任何资料记录！</font></p>
        </td>
      </tr>      <%
else
const maxperpage=60 '定义每一页显示的数据记录的常量
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
        <td width="63" bgcolor="#FFFFFF" height="20" align="center"> <%=rs("dw")%></a></td>
        <td width="53" bgcolor="#FFFFFF" height="20" align="center"><%=rs("syr")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="69"><%=rs("hcmc")%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="116"><%=rs("xh")%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="48"><%=rs("qssl")%></td>
        <td width="71" bgcolor="#FFFFFF" height="20" align="center">　<%=rs("sqrq")%></td>
        <td width="73" bgcolor="#FFFFFF" height="20" align="center"><%=rs("qsrq")%>　</td>
        <td width="47" bgcolor="#FFFFFF" height="20" align="center"><%=rs("je")%>　</td>
        <%if session("admin")<>"" then%>
        <%end if%>
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
  <table border="0" width="750" align="center" cellspacing="0" cellpadding="0">
    <tr>
    <td width="464">页数：<%=currentpage%>/<% =n%>
   	<%k=currentpage                                                                 
   	if k<>1 then%>
   	[<a href="fsmx1.asp?pageid=1">首页</a>]                                                                 
   	[<a href="fsmx1.asp?pageid=<%=k-1%>">上一页</a>]                                                                 
   	<%else%>
   	[首页]&nbsp;[上一页]                                                                 
   	<%end if%>
   	<%if k<>n then%>                                                                 
   	[<a href="fsmx1.asp?pageid=<%=k+1%>">下一页</a>]                                                                 
   	[<a href="fsmx1.asp?pageid=<%=n%>">尾页</a>]                                                                 
   	<%else%>
   	[下一页]&nbsp;[尾页]                                                                 
   	<%end if%>
        共有<font color="red"><%=totalput%></font>条记录                                                                
      <b><font color="#0000FF"><font size="2">&nbsp;</font></font></b></td> 
          <form action="fsmx.asp" name="pfrm">     
      <td height="23" width="282">  
        <p align="right">
        <select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 23; width: 72; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">  
          <%                                             
          i=1
          do while not i > n
          %>
          <option value="fsmx1.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>第 
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
    <%end if%>        
  </table>        
</div>        
  <p>
        　
  </p>
</center>
</body>       
       
</html>       
</script>









































































































































