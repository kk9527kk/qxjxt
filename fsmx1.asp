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
  fjgs="�����ع���˰�������Ĳ������嵥"
elseif aa="" and dw1<>"" then
  fjgs="�����ع���˰��"&dw&"������Ĳ������嵥"
elseif aa<>"" and dw1="" then
  aa1=mid(aa,2,len(aa)-2)
  fjgs="�����ع���˰��"&aa1&"�����嵥"
elseif aa<>"" and dw1<>"" then
    aa1=mid(aa,2,len(aa)-2)
  fjgs="�����ع���˰���"&dw&aa1&"�����嵥"
end if
    response.cookies("rqq") = rqq
    response.cookies("rqz") = rqz
%>

<html>
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
<title>�����ع���˰��ּ�����ʲ�����</title>
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
          </font><font color="#000080" face="����" size="6"><%=fjgs%></font></b></p>
                    <p align="center"><font face="����" size="2"><b>ͳ��ʱ�䣺��<%=rqq1%>��<%=rqz1%></b></font></p>
          <p align="right"><font color="#000080" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;   
          <a href="cxquit.asp">������ҳ</a> &nbsp;&nbsp;&nbsp;</b><b> </b></font></p>    
        </td>
	</tr>
</table>
<div align="left">
    <table border="0" width="765"  align="center" cellspacing="1" cellpadding="0" bgcolor="#000000">
      <tr> 
        <td width="63" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center"> 
          <p align="center">��λ 
        </td>
        <td width="53" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ʹ����</td>
        <td width="69" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">�Ĳ�����</td>
        <td width="116" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        �ͺ�</td>
		<td width="48" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ����</td>

        <td width="71" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        <font size="2">��������</font></td>

        <td width="73" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ��������</td>

        <td width="47" bgcolor="#FFFFFF" height="23" background="gdzc/images/bg1.gif" align="center">
        ���</td>

        
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
          <p align="center"><font color="#FF0000">��ʱû���κ����ϼ�¼��</font></p>
        </td>
      </tr>      <%
else
const maxperpage=60 '����ÿһҳ��ʾ�����ݼ�¼�ĳ���
dim currentpage '���嵱ǰҳ�ı���
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
'�������currentpage���������Ͳ�����ֵ��
'��1��������currentpage
if not isnumeric(currentpage) then
currentpage=1
end if
dim totalput,n '�������
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
        <td width="71" bgcolor="#FFFFFF" height="20" align="center">��<%=rs("sqrq")%></td>
        <td width="73" bgcolor="#FFFFFF" height="20" align="center"><%=rs("qsrq")%>��</td>
        <td width="47" bgcolor="#FFFFFF" height="20" align="center"><%=rs("je")%>��</td>
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
    <td width="464">ҳ����<%=currentpage%>/<% =n%>
   	<%k=currentpage                                                                 
   	if k<>1 then%>
   	[<a href="fsmx1.asp?pageid=1">��ҳ</a>]                                                                 
   	[<a href="fsmx1.asp?pageid=<%=k-1%>">��һҳ</a>]                                                                 
   	<%else%>
   	[��ҳ]&nbsp;[��һҳ]                                                                 
   	<%end if%>
   	<%if k<>n then%>                                                                 
   	[<a href="fsmx1.asp?pageid=<%=k+1%>">��һҳ</a>]                                                                 
   	[<a href="fsmx1.asp?pageid=<%=n%>">βҳ</a>]                                                                 
   	<%else%>
   	[��һҳ]&nbsp;[βҳ]                                                                 
   	<%end if%>
        ����<font color="red"><%=totalput%></font>����¼                                                                
      <b><font color="#0000FF"><font size="2">&nbsp;</font></font></b></td> 
          <form action="fsmx.asp" name="pfrm">     
      <td height="23" width="282">  
        <p align="right">
        <select onChange="pagei()" size="1" name="pag" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 23; width: 72; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">  
          <%                                             
          i=1
          do while not i > n
          %>
          <option value="fsmx1.asp?pageid=<%=i%>" <%if i=currentpage then %>selected<%end if%>>�� 
          <%=i%> ҳ</option>
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
        ��
  </p>
</center>
</body>       
       
</html>       
</script>









































































































































