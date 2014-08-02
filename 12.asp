

<head>
<style><!-- 
.fo{boder:0}; 
--> 
</style> 
<script language="javascript"> 
<!-- 
function tim(){ 
var n=new Date(); 
year=n.getYear()+"Äê"£» 
month=(n.getMonth()+1)+"ÔÂ"; 
day=n.getDate()+"ÈÕ"; 
hour=n.getHours(); 
min=n.getMinutes(); 
sec=n.getSeconds(); 
if(hour<10){hour="0"+hour;} 
if(min<10){min="0"+min;} 
if(sec<10){sec="0"+sec;} 
mm=year+month+day+" "+hour+":"+min+":"+sec 
document.form1.t1.value=mm; 
setInterval("tim()",100); 
} 
--> 
</script> 
</head> 
<body onload="tim();"> 
<form name="t1"> 
<input type="text" size="20" class="fo" name="t1"> 
</form> 
</body>
