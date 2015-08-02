<head>
<script language="javascript">
var Maxr=8,Maxc=12;
var imagedb=new Array(Maxr*Maxc);
for(var i=0;i<Maxr;i++){
	for(var j=0;j<Maxc;j++){
	imagedb[i*Maxc+j]=new Image(36,36);
	imagedb[i*Maxc+j].src='../images/land/land_r'+(i+1)+'_c'+(j+1)+'.gif';
	}}
imagedb[Maxc*Maxr]=new Image(36,36);
imagedb[Maxc*Maxr].src='../images/land/border.gif';
function gotorc(r,c){
	var mapsn=0;
	for(i=0;i<3;i++){
		for(j=0;j<3;j++){
			mapsn=(r-1+i)*Maxc+(c-1+j)
				if((r-1+i<0)|(c-1+j<0)|(r-1+i>=Maxr)|(c-1+j>=Maxc) ){mapsn=Maxc*Maxr;}
				document.images[i*3+j].src=imagedb[mapsn].src;
			}
		}
	}
function init(r,c){
	parent.confrm.location.replace('condition.asp');
	parent.behfrm.location.href='action.asp';
	parent.msgfrm.document.open();
	parent.msgfrm.document.writeln("<head><link rel=\'stylesheet\' href=\'..\/style.css\'><\script language=javascript>function Autoscroll(){self.window.scroll(0,65000);setTimeout('Autoscroll();',1000);}Autoscroll();<\/script><\/head><body><font color=ff0000>【浏览器涮新】欢迎来到探险世界<br></font>");
	gotorc(r,c);
}	
</script>
<style>
#Layer0 {POSITION: absolute;HEIGHT: 36px;WIDTH: 36px; Z-INDEX:3;left:36px;top=36px}
</style>
</head>
<body topmargin="0" leftmargin="0" onload="init(<%=session("Ba_jxqy_userposr")%>,<%=session("Ba_jxqy_userposc")%>)" oncontextmenu="self.event.returnValue=false">
<table border="0" cellspacing="0" cellpadding="0">
<tr>
	<td><img name="landtl" src="../images/land/border.gif" border="0" WIDTH="36" HEIGHT="36"></td>
	<td><img name="landt" src="../images/land/border.gif" border="0" WIDTH="36" HEIGHT="36"></td>
	<td><img name="landtr" src="../images/land/border.gif" border="0" WIDTH="36" HEIGHT="36"></td>
</tr>
<tr>
	<td><img name="landl" src="../images/land/border.gif" border="0" WIDTH="36" HEIGHT="36"></td>
	<td><img name="land" src="../images/land/border.gif" border="0" WIDTH="36" HEIGHT="36"></td>
	<td><img name="landr" src="../images/land/border.gif" border="0" WIDTH="36" HEIGHT="36"></td>
</tr>
<tr>
	<td><img name="landbl" src="../images/land/border.gif" border="0" WIDTH="36" HEIGHT="36"></td>
	<td><img name="landb" src="../images/land/border.gif" border="0" WIDTH="36" HEIGHT="36"></td>
	<td><img name="landbr" src="../images/land/border.gif" border="0" WIDTH="36" HEIGHT="36"></td>
</tr>
</table>
<div id="Layer0"><img src="../images/horse.gif" WIDTH="38" HEIGHT="34"></div>
</body>
