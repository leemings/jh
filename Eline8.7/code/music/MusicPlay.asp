<!--#include file="conn.asp"-->
<%
if request("id")="" then
		response.write"<SCRIPT language=JavaScript>alert('哈哈! 出错啦! 您未选择歌曲! ');"
		response.write"javascript:window.close();</SCRIPT>"
	else
end if
conn.close
set conn=nothing
Wma=replace(request("id"),"","")
Wma=Wma
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>一线视听</title>
<script id="chromeless" src="images/TT78_Play/pz_chromeless_.jpg"></script>
<script language="JavaScript">
<!--
function openChromeslessWindow(openUrl, winName, wWidth, wHeight, wPosx, wPosy, wTIT,
windowBORDERCOLOR, windowBORDERCOLORsel, windowTITBGCOLOR, windowTITBGCOLORsel,
bCenter, sFontFamily, sFontSize, sFontColor){
openchromeless(openUrl,winName, wWidth, wHeight, wPosx, wPosy, wTIT, wTIT ,
windowBORDERCOLOR, windowBORDERCOLORsel, windowTITBGCOLOR, windowTITBGCOLORsel,
bCenter, sFontFamily, sFontSize, sFontColor);
}
//-->
</script></head><meta http-equiv="refresh" content="3;URL=/"><body onload="openChromeslessWindow('Yxplay_.asp?id=<%=Wma%>','yajule',298,230,228,158,'一线视听', '#000000', '#000000', '#999988', '#808040' ,false,'Arial, Helvetica, sans-serif', '1','#000000');"></body></html>
