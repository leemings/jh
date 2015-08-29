<!--#include file="conn.asp"-->
<%
if request("checked")="" then
		response.write"<SCRIPT language=JavaScript>alert('哈哈! 出错啦! 至少要选择歌曲一首以上才能连播!');"
		response.write"javascript:window.close();</SCRIPT>"
	else
end if
conn.close
set conn=nothing
Wma=replace(request("checked"),"","")
Wma=Wma
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="www.T365.Net">
<title>蓝雨听吧</title>
<script>
<!--
document.write(unescape("%3Cscript%3E%0D%0A%3C%21--%0D%0Adocument.write%28unescape%28%22%253Cscript%253E%250D%250A%253C%2521--%250D%250Adocument.write%2528unescape%2528%2522%25253Cscript%252520id%25253D%252522chromeless%252522%252520src%25253D%252522image/pz_chromeless_.jpg%252522%25253E%25253C/script%25253E%25250D%25250A%25253Cscript%252520language%25253D%252522JavaScript%252522%25253E%25250D%25250A%25253C%252521--%25250D%25250Afunction%252520openChromeslessWindow%252528openUrl%25252C%252520winName%25252C%252520wWidth%25252C%252520wHeight%25252C%252520wPosx%25252C%252520wPosy%25252C%252520wTIT%25252C%25250D%25250AwindowBORDERCOLOR%25252C%252520windowBORDERCOLORsel%25252C%252520windowTITBGCOLOR%25252C%252520windowTITBGCOLORsel%25252C%25250D%25250AbCenter%25252C%252520sFontFamily%25252C%252520sFontSize%25252C%252520sFontColor%252529%25257B%25250D%25250Aopenchromeless%252528openUrl%25252CwinName%25252C%252520wWidth%25252C%252520wHeight%25252C%252520wPosx%25252C%252520wPosy%25252C%252520wTIT%25252C%252520wTIT%252520%25252C%25250D%25250AwindowBORDERCOLOR%25252C%252520windowBORDERCOLORsel%25252C%252520windowTITBGCOLOR%25252C%252520windowTITBGCOLORsel%25252C%25250D%25250AbCenter%25252C%252520sFontFamily%25252C%252520sFontSize%25252C%252520sFontColor%252529%25253B%25250D%25250A%25257D%25250D%25250A//--%25253E%25250D%25250A%25253C/script%25253E%2522%2529%2529%253B%250D%250A//--%253E%250D%250A%253C/script%253E%22%29%29%3B%0D%0A//--%3E%0D%0A%3C/script%3E"));
//-->
</script>
</head>
<body onload="window.open('Yxplay_.asp?id=<%=Wma%>','nbsky',128,158,false,'Arial, Helvetica, sans-serif', '2','#000000');setTimeout(window.close, 500);">
<!--webbot BOT="HTMLMarkup" endspan --></body></html>