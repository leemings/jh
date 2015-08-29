<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then 
	Response.Write "<script Language=Javascript>top.location.href='http://www.51eline.com';alert('提示：对不起，您还没有登陆江湖！');</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：想听歌请先进入江湖聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 银两,金币 FROM 用户 WHERE 姓名='"& session("sjjh_name") &"'",conn,1,3
if rs("银两")<100000 or rs("金币")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "身上没100000银两或金币不足1个，不能连播！"
	response.end
end if
rs("银两")=rs("银两")-100000
rs("金币")=rs("金币")-1
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="conn.asp"-->
<%
if request("checked")="" then
		response.write"<SCRIPT language=JavaScript>alert('哈哈！出错啦！至少要选择歌曲一首以上才能连播！');"
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
<meta name="author" content="www.hejiabo.com">
<title>一线视听 ← wWw.51Eline.COM</title>
<script>
<!--
document.write(unescape("%3Cscript%3E%0D%0A%3C%21--%0D%0Adocument.write%28unescape%28%22%253Cscript%253E%250D%250A%253C%2521--%250D%250Adocument.write%2528unescape%2528%2522%25253Cscript%252520id%25253D%252522chromeless%252522%252520src%25253D%252522image/pz_chromeless_.jpg%252522%25253E%25253C/script%25253E%25250D%25250A%25253Cscript%252520language%25253D%252522JavaScript%252522%25253E%25250D%25250A%25253C%252521--%25250D%25250Afunction%252520openChromeslessWindow%252528openUrl%25252C%252520winName%25252C%252520wWidth%25252C%252520wHeight%25252C%252520wPosx%25252C%252520wPosy%25252C%252520wTIT%25252C%25250D%25250AwindowBORDERCOLOR%25252C%252520windowBORDERCOLORsel%25252C%252520windowTITBGCOLOR%25252C%252520windowTITBGCOLORsel%25252C%25250D%25250AbCenter%25252C%252520sFontFamily%25252C%252520sFontSize%25252C%252520sFontColor%252529%25257B%25250D%25250Aopenchromeless%252528openUrl%25252CwinName%25252C%252520wWidth%25252C%252520wHeight%25252C%252520wPosx%25252C%252520wPosy%25252C%252520wTIT%25252C%252520wTIT%252520%25252C%25250D%25250AwindowBORDERCOLOR%25252C%252520windowBORDERCOLORsel%25252C%252520windowTITBGCOLOR%25252C%252520windowTITBGCOLORsel%25252C%25250D%25250AbCenter%25252C%252520sFontFamily%25252C%252520sFontSize%25252C%252520sFontColor%252529%25253B%25250D%25250A%25257D%25250D%25250A//--%25253E%25250D%25250A%25253C/script%25253E%2522%2529%2529%253B%250D%250A//--%253E%250D%250A%253C/script%253E%22%29%29%3B%0D%0A//--%3E%0D%0A%3C/script%3E"));
//-->
</script>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<body onload="MM_openBrWindow('Yxplay_.asp?id=<%=Wma%>','','width=300,height=300 top=100 left=100');setTimeout(window.close, 1000);" oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
</body></html>