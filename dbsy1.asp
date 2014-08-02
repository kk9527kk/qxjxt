<!--#include file="conn.asp" -->


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
</head>

<body topmargin="2" background="images/header_1.gif" leftmargin="0">

<table border="0" id="table1"  align="center" cellspacing="0" cellpadding="0" width="477">
	<tr>
		<td width="475">
          <p align="center"><font color="#0000FF" face="黑体" size="6">计算机耗材及配件待办事项</font>
          <p align="center">　
        </td>
	</tr>
	<tr>
		<td width="475">
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="547"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="84" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center"> 
          <p align="center">操作员 
        </td>
        <td width="84" bgcolor="#FFFFFF" height="23" background="images/bg1.gif" align="center">
        待办数量</td>
       
      </tr>
     

      <%dim mc
      sql="select * from dbsy"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,connstr,3,2
do while not rs2.eof 
rs2.delete
rs2.update
rs2.movenext
    loop
     rs2.close
   set rs2=nothing


    sql="select *  from SQSPB WHERE ZT='1' or ZT='2' or ZT='3' order by SQRQ,id "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
   do while  not rs.eof
   DW1=rs("dw")
   zt1=rs("zt")
   if zt1="1" then
       sql="select zg  from djb  WHERE dw='"&dw1&"' and  Zw='2' "
          set rs1=server.createobject("adodb.recordset")
          rs1.open sql,connstr,1,1
          if not rs1.eof then
              zg1=rs1("zg")
           end if 
   end if  
   if zt1="2" then 
       sql="select zg  from djb  WHERE   Zw='3' "
          set rs1=server.createobject("adodb.recordset")
          rs1.open sql,connstr,1,1
          if not rs1.eof then
              zg1=rs1("zg")
           end if 
   end if  
   if zt1="3" then 
   sql="select zg  from djb  WHERE   Zw='4' "
          set rs1=server.createobject("adodb.recordset")
          rs1.open sql,connstr,1,1
          if not rs1.eof then
              zg1=rs1("zg")
           end if 
   end if 
 if zg1<>"" then   
 sql="select * from dbsy"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,connstr,3,2
rs2.addnew
rs2("zg")=zg1
RS2("sl")=1
rs2.update
rs2.close
set rs2=nothing  
end if
  rs.movenext
    loop
 sql="select zg,sum(sl) as sl  from dbsy  group by zg  "
          set rs3=server.createobject("adodb.recordset")
          rs3.open sql,connstr,1,1
         do while not rs3.eof
              

  %> 
    
    <tr> 
        <td width="84" bgcolor="#FFFFFF" height="20" align="center"> <%=rs3("zg")%></td>
        <td width="84" bgcolor="#FFFFFF" height="20" align="center"><%=rs3("sl")%></td>
       <%rs3.movenext
       loop%>
      </tr>
    </table>
</div>
  <center>
  <p>
        　
  </p>
</center>
</body>       
       
</html>       
</script>












