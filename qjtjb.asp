<!--#include file="conn.asp"-->
<html>
<%
rqq=request.form("T1")
rqz=request.form("T2")
dw1=request.form("dw")
qjlb1=request.form("qjlb")

if not (isdate(rqq) and  isdate(rqz)) then
   response.write "<script language=JavaScript>{window.alert('��������ȷ�����ڸ�ʽ!');window.location.href='QJtj.asp';}</script>"
   response.end
end if 

    response.cookies("rqq") = rqq
    response.cookies("rqz") = rqz
    response.cookies("dw1") = dw1
    response.cookies("qjlb1") = qjlb1    

if DateDiff("d",rqq,rqz)<0 then 
   response.write "<script language=JavaScript>{window.alert('��������ʱ���������ʱ��ֹʱ�䣡');window.location.href='QJtj.asp';}</script>"
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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����ع���˰���������ͳ�Ʊ�</title>
<script language="javascript">
function pagei(){
window.location.href=document.pfrm.pag.value
}
</script>



</head>

<body topmargin="2" background="gdzc/header_1.gif" leftmargin="0">
<table border="0" id="table1" align="center"  cellspacing="0" cellpadding="0" width="815">
	<tr>
		<td width="813">
          <p align="center"><font color="#0000FF" face="����" size="6">�����ع���˰���<%=dw1%>������ͳ�Ʊ�</font>
        </td>
	</tr>
	<tr>
		<td width="813">
<p align="center"><b>ͳ��ʱ�䣺��<%=rqq1%>��<%=rqz1%>        </b>        </p>
        </td>
	</tr>
</table>
<table border="0" id="table1" align="center" cellspacing="0" cellpadding="0" width="819">
  <tr>
    <td width="817">
      <p align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="qjtjb1.asp"><font color="#000080" size="2">:::���ݼ�������ͳ��:::</font></a>&nbsp;&nbsp;&nbsp;&nbsp;</p>            
    </td>
  </tr>
</table>
<div align="left">
    <table border="0" width="743"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
         
        <td width="33.33%" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">ְ������</td>
        <td width="33.33%" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">�������</td>
        <td width="33.33%" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ʵ������</td>

             </tr>
             
             
             <% if dw1=""  and qjlb1="" then
                 bds="(QJzt='5' OR   QJzt='6' OR QJzt='7')  and sqsj between #"&rqq&"#"&" and #"&rqz&"#"  
                elseif  dw1=""  and qjlb1<>"" then
                 bds="(QJzt='5' OR   QJzt='6' OR QJzt='7')  AND QJLB='"&QJLB1&"' and sqsj between #"&rqq&"#"&" and #"&rqz&"#"  
                elseif dw1<>"" and qjlb1<>"" then
                  bds="(QJzt='5' OR   QJzt='6' OR QJzt='7')  and  dw='"&dw1&"' AND QJLB='"&QJLB1&"' and sqsj between #"&rqq&"#"&" and #"&rqz&"#"  
                else
                 bds="(QJzt='5' OR   QJzt='6' OR QJzt='7')  and  dw='"&dw1&"' and sqsj between #"&rqq&"#"&" and #"&rqz&"#"  
                end if
          
           sql="select sum(TS),sum(sxts) from QXJDJB  where "&bds  
           set rs=server.createobject("adodb.recordset")
           rs.open sql,connstr,1,1
           sl=rs(0)
           sl1=rs(1)
          rs.close

     %>

    
      <tr> 
        <td width="33.33%" bgcolor="#FFFFFF" height="20" align="center">�ϼ�</td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="33.33%"><%=sl%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="33.33%"><%=sl1%>��</td>
     </tr>
      <%if dw1="" then
               sql="select *  from zg  order by XH  "
             ELSE
                sql="select *  from zg  where dw='"&dw1&"' order by XH "
             END IF

             
      Arr="" 
    'sql="select *  from zg  order by id where "&bds
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connstr,1,1
    
Do While Not Rs.Eof 
If Arr="" Then 
Arr=Rs("zg") 
Else 
Arr=Arr&","&Rs("zg") 
End If 
Rs.Movenext 
loop
 rs.close
Arr=split(Arr,",") 
  %> 
    <% for   n=0   to   UBound(arr)   %>        
          <% 
          
                if dw1=""  and qjlb1="" then
                 bds="(QJzt='5' OR   QJzt='6' OR QJzt='7')  and sqsj between #"&rqq&"#"&" and #"&rqz&"#"&"  and zg='"&arr(n)&"'" 

                elseif  dw1=""  and qjlb1<>"" then
                 bds="(QJzt='5' OR   QJzt='6' OR QJzt='7')  AND QJLB='"&QJLB1&"' and sqsj between #"&rqq&"#"&" and #"&rqz&"#"&"  and zg='"&arr(n)&"'" 

                elseif dw1<>"" and qjlb1<>"" then
                  bds="(QJzt='5' OR   QJzt='6' OR QJzt='7')  and  dw='"&dw1&"' AND QJLB='"&QJLB1&"' and sqsj between #"&rqq&"#"&" and #"&rqz&"#"&"  and zg='"&arr(n)&"'" 

                else
                 bds="(QJzt='5' OR   QJzt='6' OR QJzt='7')  and  dw='"&dw1&"' and sqsj between #"&rqq&"#"&" and #"&rqz&"#"&"  and zg='"&arr(n)&"'" 

                end if
           sql="select sum(TS),sum(sxts) from QXJDJB  where "&bds  
           set rs1=server.createobject("adodb.recordset")
           rs1.open sql,connstr,1,1
           sl=rs1(0)
           sl1=rs1(1)
          rs1.close
          %>
    
      <tr> 
        <td width="33.33%" bgcolor="#FFFFFF" height="20" align="center">��<a href="zgcxcl.asp?zg='<%=arr(n)%>'"><%=arr(n)%></td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="33.33%"><%=sl%>��</td>
        <td bgcolor="#FFFFFF" height="20" align="center" width="33.33%"><%=sl1%>��</td>
     </tr>
      
      
        

          
          <%next%>
   
    
    </table>
</div>
<center>
<p></p>
</center>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                           
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
�ر�˵������������������ʱ��Ϊ���ݡ��������������ͬ���ݼٵ���������ʵ���������˽̿�����¼���������</p> 
</body>       
       
</html>       
</script>