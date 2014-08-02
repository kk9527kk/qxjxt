<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='login.asp';}</script>"
response.end
end if
lid=request.QueryString("id")
uid=request.Cookies("adminuser")
'response.write(uid)
sql="select * from qxjdjb  where zg='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
    dw1=rs1("dw")
    zg1=rs1("zg")
    zw1=rs1("zw")
    
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
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局计算机耗材及配件申请审批表</title>
</head>

<body background="images/header_1.gif">
<p align="center"><font color="#0000FF" face="黑体" size="6">云阳县国家税务请销假查询<br>
</font><font color="#000080" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></p>

<table border="1" align="center" cellpadding="0" cellspacing="0" width="64%" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="46">
  <tr>
    <td width="100%" height="30" background="images/bg1.gif">
      <p align="center">根据给定的条件进行查询</td>
  </tr>
  <tr>
    <td width="100%" height="12">
    <form method="POST" action="cxcl.asp"  name="form1">
           <p align="center">　 </p>
           <p align="center">查询时间起：  <input name="T1" type="text" onClick="WdatePicker()">&nbsp;&nbsp;               
           查询时间止：  <input name="T2" type="text" onClick="WdatePicker()"> </p>               
        <script language="javascript" type="text/javascript" src="rq/My97DatePicker/WdatePicker.js"></script>
           <p align="center">所属单位：  
           <select   size="1"   name="dw"   onchange   ="changeselect(document.form1.dw.options[document.form1.dw.selectedIndex].value)">   
                              <%=DbCombox()%></select> 
          
           &nbsp;&nbsp; 职工姓名： <select size="1" name="zg">              
           </select> </p>
           　<p align="center"><input type="submit" value="确认" name="B1"><input type="reset" value="重写" name="B2"></p></td>  </form>
  </tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="64%" height="55">

<%   
  function   DbCombox()   
    dim   rs,sql,msg   
    sql   =   "select   *   from   dw"   
    set   rs   =   conn.execute(sql)   
    while   not   rs.eof   
      msg   =   msg   &   "<option   value="""   &   rs("dw")   &   """>"   &   rs("dw")   &   "</option>"   
      rs.movenext   
    wend   
    rs.close   
    set   rs   =   nothing   
    DbCombox   =   msg   
  End   function   
  %>  
  <%lid=request.QueryString("id")
sql="select * from qxjdjb where zg='"&uid&"'"
set editrs=server.createobject("adodb.recordset")
editrs.open sql,connstr,3,2
'dw=editrs("dw")
'zg=editrs("zg")
%>
<table>



  <tr>
    <td width="100%" height="55">
    </td>
  </tr>
</body>
<script   language   ="javascript"   >   
  Citys   =   new   Array();   
  <%   
  dim   rs,sql,i   
  sql   =   "select   *   from   zg"   
  set   rs   =   Conn.execute(sql)   
  i   =   0   
  while   not   rs.eof   
  %>   
  Citys[<%=i%>]   =new   Array("<%=rs("dw")%>","<%=rs("zg")%>");   
  <%   
  i   =   i   +   1   
  rs.movenext   
  wend   
  rs.close   
  set   rs   =   nothing   
  %>   
    
  function   changeselect(selvalue){   
    var   selvalue   =   selvalue;   
    var   i;   
    document.form1.zg.length   =   0   ;   
    document.form1.zg.options[document.form1.zg.length]   =   new   Option("请选择","");   
    for   (i   =   0   ;i   <Citys.length;i++){   
      if(Citys[i][0]==selvalue){   
        document.form1.zg.options[document.form1.zg   .length]   =   new   Option(Citys[i][1],Citys[i][1]);   
      }   
    }   
  }   
    
  document.form1.zg.options[document.form1.zg.length]   =   new   Option("请选择","");   
    
  </script>   



