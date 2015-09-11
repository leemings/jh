<!--#include file="conn.asp"-->
<%if session("aqjh_name")="" then
response.redirect "../login.asp"
response.end
else
sql= "select * from 股票 "
set rss=conn.execute(sql)
if rss("状态")="开" then 
DO while not RSS.EOF   
sql="update 股票 set 状态='封' ,交易量=0 "
conn.execute sql
rss.movenext
loop

end if


%>
<html>
<head>
<title>收市了</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFEF4" text="#000000">
<div align="center">
  <p>对不起，股票市场已经收市,请于每天的 ＡＭ 5:00～ＰＭ 24:00 再来。 </p>
  <p><OBJECT  id=closes  type="application/x-oleobject"  classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11" width="14" height="14">   
<param  name="Command"  value="Close">   
</object> <a href="#" onclick="closes.Click()">关闭本页 </a>
</p>
</div>
</body>
</html>
<%
end if
CloseDatabase		'关闭数据库

%>