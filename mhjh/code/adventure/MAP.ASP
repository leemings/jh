<head>
<script language="javascript">
var Maxr=14,Maxc=21;
var imagedb=new Array(Maxr*Maxc);
for(var i=0;i<Maxr;i++){
for(var j=0;j<Maxc;j++){
imagedb[i*Maxc+j]=new Image(35,23);
imagedb[i*Maxc+j].src='../image/land/land_r'+(i+1)+'_c'+(j+1)+'.jpg';
}}
imagedb[Maxc*Maxr]=new Image(35,23);
imagedb[Maxc*Maxr].src='../image/land/border.gif';
function gotorc(r,c){
var mapsn=0;
for(i=0;i<9;i++){
for(j=0;j<18;j++){
mapsn=(r-1+i)*Maxc+(c-1+j)
if((r-1+i<0)|(c-1+j<0)|(r-1+i>=Maxr)|(c-1+j>=Maxc) ){mapsn=Maxc*Maxr;}
document.images[i*18+j].src=imagedb[mapsn].src;
}
}
}
function init(r,c){
parent.confrm.location.replace('condition.asp');
parent.behfrm.location.href='action.asp';
parent.msgfrm.document.open();
parent.msgfrm.document.writeln("<head><link rel=\'stylesheet\' href=\'../../style.css\'><\script language=javascript>function Autoscroll(){self.window.scroll(0,65000);setTimeout('Autoscroll();',1000);}Autoscroll();<\/script><\/head><body><font color=ff0000>【浏览器涮新】欢迎来到探险世界<br></font>");
gotorc(r,c);
}
</script>
<style>
#Layer0 {POSITION: absolute;HEIGHT:23px;WIDTH:23px; Z-INDEX:18;left:23px;top=23px}
</style>
</head>
<body topmargin="0" leftmargin="0" onload="init(<%=session("yx8_mhjh_userposr")%>,<%=session("yx8_mhjh_userposc")%>)" oncontextmenu="self.event.returnValue=false">
<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
</tr>
<tr>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
</tr>
<tr>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
</tr>
<tr>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
</tr>
<tr>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
</tr>
<tr>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
</tr>
<tr>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landt" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landtr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
</tr>
<tr>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="land" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
</tr>
<tr>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbl" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landb" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
<td><img name="landbr" src="../image/land/border.gif" border="0" WIDTH="35" HEIGHT="23"></td>
</tr>
</table>
<div id="Layer0"><img src="004.gif" WIDTH="50" HEIGHT="50"></div>
</body>