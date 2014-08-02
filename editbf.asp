<%if session("admin")="" then
response.redirect "index.asp"
end if%>
<!--#include   file="Conn.asp"-->   
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
sql="select * from kp where id="&lid
set editrs=server.createobject("adodb.recordset")
editrs.open sql,connstr,3,2
dw=editrs("dw")
zg=editrs("zg")
zcmc=editrs("zcmc")
xh=editrs("xh")
xlh=editrs("xlh")
pzqk=editrs("pzqk")
je=editrs("je")
djrq=date()
%>
<html>
<script   language   ="javascript"   >   
  Citys   =   new   Array();   
  <%   
  dim   rs,sql,i   
  sql   =   "select   *   from   djb"   
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
<title>云阳县国家税务局计算机资产登记维护</title>
</head>

<body topmargin="0" background="txl/header_1.gif">

<div align="center"> 　</div>
<div align="center">
  <center>
  <table border="0" width="279">
    <tr>
      <td width="353" height="8"></td>
    </tr>
    <tr>
      <td width="353">
        <div align="left">
          <table border="0" width="268" cellspacing="1" cellpadding="0" bgcolor="#000000" height="256">
         
 <form action="save1.asp?id=<%=lid%> "  method="post"   name="form1" >
            <tr>
              <td width="281" background="images/bg1.gif" height="21">
                <p align="center"><b>登记信息修改</b></td>
            </tr>
          </center>
            <tr>
              <td width="281" bgcolor="#FFFFFF" height="231">
                <table border="0" width="261" height="169">
                  <tr>
                    <td width="76" valign="bottom" height="20"> 
                      <p align="right">&nbsp;使用单位： 
                    </td>
                    <td width="171" height="20"> 
						<select   size="1"   name="dw"   onchange   ="changeselect(document.form1.dw.options[document.form1.dw.selectedIndex].value)">   
                              <%=DbCombox()%></select>                      </td>
                  </tr>
                  <tr> 
                    <td width="30%" valign="bottom" height="20" align="right"> 
                      <p align="right">姓 名： 
                    </td>
                    <td width="70%" height="20"> 
                                           
                <select   size="1"   name="zg"></select>                          
                    </td>
                  </tr>
                  <tr> 
                    <td width="76" align="right" valign="bottom" height="20">资产名称：</td>
                    <td width="171" height="20"> 
                      <select size="1" name="zcmc">
                         <option value="<%=zcmc%>"><%=zcmc%></option>
                         <option value="计算机">计算机</option>
                          <option value="激光打">激光打</option>
                          <option value="平推打">平推打</option>
			<option value="普针打">普针打</option>
		         <option value="显示器">显示器</option>
                         <option value="扫描仪">扫描仪</option>
                         <option value="移动盘">移动盘</option>
                         <option value="笔记本">笔记本</option>
			<option value="服务器">服务器</option>
                          <option value="交换机">交换机</option>
                         
							</select>
                    </td>
                  </tr>
                  <tr> 
                    <td width="76" align="right" valign="bottom" height="20">型 
                      号：</td>
                    <td width="171" height="20"> 
                      <input type="text" name="xh" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%" value="<%=xh%>">
                    </td>
                  </tr>
                  <tr>
                    <td align="right" valign="bottom" height="20" width="76"> 序列号：</td>
                    <td height="20" width="171"> 
                      <input type="text" name="xlh" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%" value="<%=xlh%>"
                    </td>
                  </tr>
					<tr>
                    <td align="right" valign="bottom" height="20" width="76"> 配置情况：</td>
                    <td height="20" width="171"> 
                      <input type="text" name="pzqk" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%" value="<%=pzqk%>">
                    </td>
                  </tr>
					<tr>
                    <td align="right" valign="bottom" height="20" width="76"> 资产价格：</td>
                    <td height="20" width="171"> 
                                              
                        <input type="text" name="je" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%" value="<%=je%>">
                    </td>
                  </tr>
                  <tr> 
                    <td width="76" valign="top" height="56">
                      <p align="right">&nbsp;&nbsp;资产状态：</p>
                    </td>
                    <td width="171" height="56" valign="top"> 
                                              
                        <select size="1" name="zt">
                         <option value="待用" selected>待用</option>
                         <option value="在用" selected>在用</option>
                          <option value="送修">送修</option>
                          <option value="待报废">待报废</option>
			<option value="已失踪">已失踪</option>
		         <option value="已报废">已报废</option>
                        </select>
                        <p>　</p>
                        <p>
                        <input type="submit" value=" 修 改 " name="edit" style="background-color: #DDDDDD; background-repeat: repeat; background-attachment: scroll; font-size: 9pt; height: 20; width: 64; background-image: url('images/bg1.gif'); border: 1px groove #000000; background-position: 0% 50%"><input type="reset" value=" 取 消 " name="reset" style="background-color: #DDDDDD; background-repeat: repeat; background-attachment: scroll; font-size: 9pt; height: 20; width: 71; background-image: url('images/bg1.gif'); border: 1px groove #000000; background-position: 0% 50%"> 
                      </form>        
                        </p>
                    </td>
                  </tr>
                  <tr> 
                    <td width="253" colspan="2" height="14"> 
                      <p align="center"> 
                        &nbsp;&nbsp;&nbsp;&nbsp;
                      </p>
                    </td>
                  </tr></form>
                </table>             
                  </td>             
            </tr>             
          </table>             
        </div>             
      </td>             
    </tr>             
  </table>             
</div>             
             
<div align="center">             
  <center>             
  <table border="0" width="84" cellspacing="0" cellpadding="0">             
    <tr>             
      <td height="8" width="84"></td>             
    </tr>             
    <tr>             
      <td width="84">        
        <p align="center"><a href="start.asp">返回首页</a></td>             
    </tr>             
    <tr>             
      <td height="8" width="84"></td>             
    </tr>             
    <tr>             
        <td height="35" valign="bottom" width="84"> 
          <p align="center">&nbsp;
        </td>              
    </tr>              
  </table>              
  </center>              
</div>              
              
</body>              
              
</html>              














