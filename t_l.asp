<html>
<head>
<title>button_title</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script>
function oa_tool1(){
if(window.parent.oa_frame.cols!="0,10,*"){
show.src="images/forward.gif"
oa_tree.title="����"}
else{
show.src="images/backward.gif"
oa_tree.title="��ʾ"}
}
function oa_tool(){
if(window.parent.oa_frame.cols=="0,10,*"){
show.src="images/forward.gif"
oa_tree.title="����"
window.parent.oa_frame.cols="160,10,*";
}
else{
show.src="images/backward.gif"
oa_tree.title="��ʾ"
window.parent.oa_frame.cols="0,10,*";}
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="return oa_tool1()">
<table width="11" border="0" height="100%" cellpadding="0" cellspacing="0" align="left">
  <tr align="center">
    <td bgcolor=#689ACE>
      <div id=oa_tree onclick="oa_tool();" title=����><img id=show src=images/forward.gif style="CURSOR: hand;"></div>
      </td>
  </tr>
</table>
</body>
</html>
