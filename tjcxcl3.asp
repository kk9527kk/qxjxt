<!--#include file="conn.asp"-->
<html>
<%
rqq=request.cookies("rqq")
rqz=request.cookies("rqz")
if not (isdate(rqq) and  isdate(rqz)) then
   response.write "<script language=JavaScript>{window.alert('��������ȷ�����ڸ�ʽ!');window.location.href='cxtj.asp';}</script>"
   response.end
end if 

    response.cookies("rqq") = rqq
    response.cookies("rqz") = rqz
if DateDiff("d",rqq,rqz)<0 then 
   response.write "<script language=JavaScript>{window.alert('��������ʱ���������ʱ��ֹʱ�䣡');window.location.href='cxtj.asp';}</script>"
   response.end
end if

rqq1=FormatDateTime(rqq,1)
rqz1=FormatDateTime(rqz,1)

%>

<head>
<STYLE type=text/css>A:link {
	COLOR: #000000; FONT-FAMILY: ����; TEXT-DECORATION: none
}
A:visited {
	COLOR: #000000; FONT-FAMILY: ����; TEXT-DECORATION: none
}
A:active {
	FONT-FAMILY: ����; TEXT-DECORATION: underline
}
A:hover {
	COLOR: #84bd6b; TEXT-DECORATION: underline overline
}
BODY {
	COLOR: #000000; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
TABLE {
	COLOR: #000000; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
.f24 {
	COLOR: #ff0000; FONT-FAMILY: ����; FONT-SIZE: 9pt; TEXT-DECORATION: underline overline
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
<title>�����ع���˰��ּ�����ʲ�ͳ�Ʊ�</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>



</head>

<body topmargin="2" background="gdzc/header_1.gif" leftmargin="0">
<table border="0" id="table1" align="center"  cellspacing="0" cellpadding="0" width="880">
	<tr>
		<td width="878">
          <p align="center"><font color="#0000FF" face="����" size="6">�����ع���˰��ּ����������Ĳ�����ͳ�Ʊ�</font>
        </td>
	</tr>
	<tr>
		<td width="878">
<p align="center"><b>ͳ��ʱ�䣺��<%=rqq1%>��<%=rqz1%>        </b>        </p>
        </td>
	</tr>
</table>
<table border="0" id="table1" align="center" cellspacing="0" cellpadding="0" width="880">
  <tr>
    <td width="878">
      <p align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="tjcxcl1.asp"><font color="#000080" size="2">:::����������ͳ��:::</font></a>&nbsp;&nbsp;&nbsp;&nbsp;</p>  
    </td>
  </tr>
</table>
<div align="left">
    <table border="0" width="884"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
         <td width="80" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ��λ</td>
        <td width="80" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">���ϼ�</td>
        <td width="80" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">����</td>
        <td width="80" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ī��</td>
		<td width="80" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ɫ��</td>

		<td width="80" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ɫ����</td>

		<td width="80" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ����</td>

		<td width="80" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        �ƶ�Ӳ��</td>

		<td width="80" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ��ӡֽ</td>

		<td width="81" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        <font size="2">�������</font></td>

		<td width="81" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ����</td>

             </tr>
             <% bds="zt='4'  and qsrq between #"&rqq&"#"&" and #"&rqz&"#"  
          sql="select sum(JE)  from sqspb  where "&bds
       '   response.write(sql)
           set rs=server.createobject("adodb.recordset")
           rs.open sql,connstr,1,1
           sl=rs(0)
          %>

             
              <% bds="zt='4'  and hcmc='����' and qsrq between #"&rqq&"#"&" and #"&rqz&"#"  
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs1=server.createobject("adodb.recordset")
           rs1.open sql,connstr,1,1
           sl1=rs1(0)
          %>

             <% bds="zt='4'  and hcmc='ī��' and qsrq between #"&rqq&"#"&" and #"&rqz&"#" 
         sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs2=server.createobject("adodb.recordset")
           rs2.open sql,connstr,1,1
           sl2=rs2(0)
          %>

        <% bds="zt='4'  and hcmc='ɫ��' and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
           sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(qssl)
           set rs3=server.createobject("adodb.recordset")
           rs3.open sql,connstr,1,1
           sl3=rs3(0)
          %>

        <% bds="zt='4'  and hcmc='ɫ����' and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs4=server.createobject("adodb.recordset")
           rs4.open sql,connstr,1,1
           sl4=rs4(0)
          %>

        <% bds="zt='4'  and hcmc='����' and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
 
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs5=server.createobject("adodb.recordset")
           rs5.open sql,connstr,1,1
           sl5=rs5(0)
          %>

        <% bds="zt='4'  and hcmc='�ƶ�Ӳ��' and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs6=server.createobject("adodb.recordset")
           rs6.open sql,connstr,1,1
           sl6=rs6(0)
          %>

        <% bds="zt='4'  and hcmc='��ӡֽ' and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs7=server.createobject("adodb.recordset")
           rs7.open sql,connstr,1,1
           sl7=rs7(0)
          %>

        <% bds="zt='4'  and hcmc='�������' and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs8=server.createobject("adodb.recordset")
           rs8.open sql,connstr,1,1
           sl8=rs8(0)
          %>

        <% bds="zt='4'  and hcmc='����' and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
 
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs9=server.createobject("adodb.recordset")
           rs9.open sql,connstr,1,1
           sl9=rs9(0)
          %>

             <tr> 
         <td width="80" bgcolor="#FFFFCC" height="23" align="center">
        <a href="fsmx.asp">�ϼ�</a></td>
        <td width="80" bgcolor="#FFFFCC" height="23" align="center"><%=sl%>��</td>
        <td width="80" bgcolor="#FFFFCC" height="23" align="center"><a href="fsmx.asp?hcmc='����'"><%=sl1%></a></td>
        <td width="80" bgcolor="#FFFFCC" height="23" align="center"><a href="fsmx.asp?hcmc='ī��'">
        <%=sl2%></a></td>
		<td width="80" bgcolor="#FFFFCC" height="23" align="center"><a href="fsmx.asp?hcmc='ɫ��'">
        <%=sl3%></a></td>

		<td width="80" bgcolor="#FFFFCC" height="23" align="center"><a href="fsmx.asp?hcmc='ɫ����'">
        <%=sl4%></a></td>

		<td width="80" bgcolor="#FFFFCC" height="23" align="center"><a href="fsmx.asp?hcmc='����'">
      <%=sl5%></a></td>

		<td width="80" bgcolor="#FFFFCC" height="23" align="center"><a href="fsmx.asp?hcmc='�ƶ�Ӳ��'">
        <%=sl6%></a></td>

		<td width="80" bgcolor="#FFFFCC" height="23" align="center"><a href="fsmx.asp?hcmc='��ӡֽ'">
       <%=sl7%></a></td>

		<td width="81" bgcolor="#FFFFCC" height="23" align="center"><a href="fsmx.asp?hcmc='�������'">
        <font size="2"><%=sl8%></font></a></td>

		<td width="81" bgcolor="#FFFFCC" height="23" align="center"><a href="fsmx.asp?hcmc='����'">
        <%=sl9%></a></td>

             </tr>
             
             <% rs.close
             rs1.close
       rs2.close
    rs3.close
rs4.close
rs5.close
rs6.close
rs7.close
rs8.close
rs9.close

     %>


      <%
      Arr="" 
    sql="select *  from dw  order by id"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    
Do While Not Rs.Eof 
If Arr="" Then 
Arr=Rs("dw") 
Else 
Arr=Arr&","&Rs("dw") 
End If 
Rs.Movenext 
loop
 rs.close
Arr=split(Arr,",") 
  %> 



     
    <% for   n=0   to   UBound(arr)   %>       
      
      <tr> 
        
        <td width="80" bgcolor="#FFFFFF" height="20" align="center"><a href="fsmx.asp?dw='<%=arr(n)%>'"><%=arr(n)%></a></td>
        
        <% bds="zt='4'   and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"  
          sql="select sum(JE)  from sqspb  where "&bds
       '   response.write(sql)
           set rs=server.createobject("adodb.recordset")
           rs.open sql,connstr,1,1
           sl=rs(0)
          %>

        <% bds="zt='4'  and hcmc='����' and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs1=server.createobject("adodb.recordset")
           rs1.open sql,connstr,1,1
           sl1=rs1(0)
          %>

        <% bds="zt='4'  and hcmc='ī��' and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs2=server.createobject("adodb.recordset")
           rs2.open sql,connstr,1,1
           sl2=rs2(0)
          %>

        <% bds="zt='4'  and hcmc='ɫ��' and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
           '   response.write(sql)
           set rs3=server.createobject("adodb.recordset")
           rs3.open sql,connstr,1,1
           sl3=rs3(0)
          %>

        <% bds="zt='4'  and hcmc='ɫ����' and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs4=server.createobject("adodb.recordset")
           rs4.open sql,connstr,1,1
           sl4=rs4(0)
          %>

        <% bds="zt='4'  and hcmc='����' and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs5=server.createobject("adodb.recordset")
           rs5.open sql,connstr,1,1
           sl5=rs5(0)
          %>

        <% bds="zt='4'  and hcmc='�ƶ�Ӳ��' and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs6=server.createobject("adodb.recordset")
           rs6.open sql,connstr,1,1
           sl6=rs6(0)
          %>

        <% bds="zt='4'  and hcmc='��ӡֽ' and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs7=server.createobject("adodb.recordset")
           rs7.open sql,connstr,1,1
           sl7=rs7(0)
          %>

        <% bds="zt='4'  and hcmc='�������' and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs8=server.createobject("adodb.recordset")
           rs8.open sql,connstr,1,1
           sl8=rs8(0)
          %>

        <% bds="zt='4' and hcmc='����' and  dw='"&arr(n)&"'" &"and qsrq between #"&rqq&"#"&" and #"&rqz&"#"
          sql="select sum(qssl)  from sqspb  where "&bds
       '   response.write(sql)
           set rs9=server.createobject("adodb.recordset")
           rs9.open sql,connstr,1,1
           sl9=rs9(0)
          %>

        
        <td bgcolor="#FFFFFF" height="20" align="center" width="80"><%=sl%>��</td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="80"><a href="fsmx.asp?hcmc='����'&dw='<%=arr(n)%>'"><%=sl1%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="80"><a href="fsmx.asp?hcmc='ī��'&dw='<%=arr(n)%>'"><%=sl2%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="80"><a href="fsmx.asp?hcmc='ɫ��'&dw='<%=arr(n)%>'"><%=sl3%></a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="80"><a href="fsmx.asp?hcmc='ɫ����'&dw='<%=arr(n)%>'"><%=sl4%>��</a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="80"><a href="fsmx.asp?hcmc='����'&dw='<%=arr(n)%>'"><%=sl5%>��</a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="80"><a href="fsmx.asp?hcmc='�ƶ�Ӳ��'&dw='<%=arr(n)%>'"><%=sl6%>��</a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="80"><a href="fsmx.asp?hcmc='��ӡֽ'&dw='<%=arr(n)%>'"><%=sl7%>��</a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="81"><a href="fsmx.asp?hcmc='�������'&dw='<%=arr(n)%>'"><%=sl8%>��</a></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="81"><a href="fsmx.asp?hcmc='����'&dw='<%=arr(n)%>'"><%=sl9%></a></td>
        <%if session("admin")<>"" then%>
        <%end if%>
        
     <% rs1.close
     rs2.close
    rs3.close
rs4.close
rs5.close
rs6.close
rs7.close
rs8.close
rs9.close
    next 
     %>
   
    
      </tr>
    </table>
</div>
<center>
</center>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      
�ر�˵���������ϼ����⣬������Ϊ����������������ǩ������Ϊ���ݡ�</p>
</body>       
       
</html>       
</script>

























































