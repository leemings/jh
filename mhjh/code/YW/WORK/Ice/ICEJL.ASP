<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.Buffer=true
if session("yx8_mhjh_username")="" then Response.Redirect "../../../error.asp?id=016"
http=Request.ServerVariables("HTTP_REFERER")
if InStr(http,"work/ice")=0 then
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>"
Response.End
end if
Response.Write "<div align='center'><font color=red>极地采冰</font></div>"
Response.Write "<br>"
Response.Write "<font color=white>完成:</font><img src='images/tong"& Session("icejl1") &".gif'></img>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write "<font color=white>采冰:</font><img src='images/tong4.gif'></img>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write " <font color=white>已采:"&Session("icets")&"桶</font>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write " <font color=white>现有："&Session("icesl")&"桶</font>"
if (Hour(time())*3600+Minute(time())*60+Second(time()))-session("cbsj")>12 or session("cbsj")="" then
session("cbsj")=(Hour(time())*3600+Minute(time())*60+Second(time()))
Session("icejl1")=Session("icejl1")+1
if Session("icejl1")=4 then
	Session("icejl1")=0
	Session("icets")=Session("icets")+1
end if
else
Response.Write "<script Language=Javascript>alert('严重警告，不允许使用采冰机变速度齿轮！');</script>"
end if
%>
<html>
<head>
<title>极地采冰</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><LINK href="../../../style2.css" rel=stylesheet>
</head>
<body bgcolor="#000000" text="#FFFFFF" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"></div>
</body>
</html>
