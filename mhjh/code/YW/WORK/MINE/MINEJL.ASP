<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.Buffer=true
if session("yx8_mhjh_username")="" then Response.Redirect "../../error.asp?id=016"
http=Request.ServerVariables("HTTP_REFERER")
if InStr(http,"work/mine")=0 then
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>"
Response.End
end if
Response.Write "<div align='center'><font color=blue></font>"
Response.Write "<font color=red>深山采矿</font></div>"
Response.Write "<div align='center'><font color=white>采矿数量:</font>"
Response.Write "<font color=white>已采:<font color=red>"&Session("minets")&"</font>块矿石</font>"
Response.Write "<font color=white>现有：<font color=red>"&Session("minesl")&"</font>块矿石</font></div>"
if (Hour(time())*3600+Minute(time())*60+Second(time()))-session("minesj")>22 or session("minesj")="" then
session("minesj")=(Hour(time())*3600+Minute(time())*60+Second(time()))
Session("minets")=Session("minets")+1
else
Response.Write "<script Language=Javascript>alert('严重警告，不允许使用鼠标控制程序或速度齿轮！');</script>"
end if
%>
<html>
<head>
<title>深山采冰</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../../style2.css" rel=stylesheet></head>
<body bgcolor="#000000" text="#FFFFFF" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center" style="width: 766; height: 211"></div>
</body>
</html>
