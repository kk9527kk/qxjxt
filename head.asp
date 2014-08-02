<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<%
uid=request.Cookies("adminuser")
sql="select * from zg  where id='"&uid&"'"
set rs1=server.createobject("adodb.recordset")
rs1.open sql,connstr,3,2
if not rs1.eof then
     zga=rs1("zg")
end if
rs1.close
%>


<head><STYLE type=text/css>A:link {
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
	COLOR: #000000; FONT-FAMILY: 宋体; FONT-SIZE: 9pt;
	background: url("images/aaa.gif") repeat-x fixed top;;
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

<link rel="StyleSheet" href="dtree.css" type="text/css" />
</head>

<body  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#8ab4fe">

		<table width="965" border="0" cellspacing="0" cellpadding="0" align=center height="1">
		<tr>			
    <td width="165" height="1" rowspan="2">
　</td>
	<td align=left width="794" height="32" colspan="2"><font color=red>&nbsp;&nbsp;<img border="0" src="images/logo.jpg" width="578" height="44">         
      </font>
      </td> 
     

		
		</tr>
		<tr>			
	<td align=left width="434" height="1"><script language="javascript">
function showtime(){
currtime.innerText=new Date().toLocaleString();
setTimeout("showtime()",1000)
}
window.attachEvent("onload",showtime)
</script>
现在时间：<b id=currtime style="color:red"></b>

      </td> 
     

		
	<td align=left width="360" height="1"><img border="0" src="images/person.GIF">登录用户：&nbsp;<%=zga%>
      </td> 
     

		
		</tr>
	</table>
	</body>



