<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>

<%



a8=2356
a7="123"
a6=(CDbl(a7)/CDbl(a8))*100

strA8=Split(CStr(a8),".")%>
<table>
      <TD align=center vAlign=middle bordercolor="#000000" bgColor=#ffffff height="21" bordercolordark="#FFFFFF" bordercolorlight="#000000"> 
        <div align="right"><font size="2"><%=FormatNumber(a8,5,0)%><P></P>
       <%=FormatNumber(a6,2,-1)%>
 </font></div></TD>
      

</body>






