<html>
<head>
<title>高炉练铁</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#000000" text="#FFFFFF" topmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"></div>
</body>
</html>
<%
http=Request.ServerVariables("HTTP_REFERER")
if InStr(http,"work/tie")=0 then
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>"
Response.End
end if
Response.Write "<div align='center'><font color=white>高炉练铁</font></div>"
Response.Write "<div align='center'><font color=white>采矿数量:</font>"
Response.Write "<font color=white>已炼:<font color=red>"& Session("tiets") &"</font>块铁</font>"
Response.Write "<font color=white>现有：<font color=red>"& Session("tiesl") &"</font>块矿石</font></div>"
if Session("tiesl")<=0 then
%>
<script language="vbscript">
  alert("您现在的矿石已经练完，请选择卖出换得银两，需要再采矿石再练铁！")
</script>
<%
response.end
end if
if (Hour(time())*3600+Minute(time())*60+Second(time()))-session("tiesj")>22 or session("tiesj")="" then
session("tiesj")=(Hour(time())*3600+Minute(time())*60+Second(time()))
Session("tiets")=Session("tiets")+1
Session("tiesl")=Session("tiesl")-2
else
Response.Write "<script Language=Javascript>alert('严重警告，不允许使用鼠标控制程序或速度齿轮！');</script>"
end if
%>
