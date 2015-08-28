<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
   Application("yx8_mhjh_hong")="无"
   Application("yx8_mhjh_hei")="无"
 msg="系统已经清除比武者"
%>
<head><title></title><LINK href="../css3.css" rel=stylesheet><script language=javascript>setTimeout('history.back();',3000);</script>
</head>
<body oncontextmenu=self.event.returnValue=false topMargin=150 background='../bg1.gif' v>
<div align=center><%=msg%><br><br><a href="anqi.asp">返回</a></div> 
</body> 
