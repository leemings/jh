<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username<>Application("yx8_mhjh_admin") then Response.Redirect "../error.asp?id=016"
   if Application("yx8_mhjh_biwu")="" then 
      Application("yx8_mhjh_biwu")="ok"
       msg="比武系统开启"
   else
      Application("yx8_mhjh_biwu")=""
      msg="比武系统关闭"
   end if   
%>
<head><title></title><LINK href="../css3.css" rel=stylesheet><script language=javascript>setTimeout('history.back();',3000);</script>
</head>
<body oncontextmenu=self.event.returnValue=false topMargin=150 background='../bg1.gif'>
<div align=center><%=msg%><br><br><a href="anqi.asp">返回</a></div> 
</body> 
