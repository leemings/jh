<!--#include file="data.asp"--> 
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
sql="select 下注者,参赛者,胜利者,数量,赔率,场次,领奖 from 投注 where 下注者='"&username&"' order by 场次"
Set Rs=connt.Execute(sql)
msg="<table border='1' cellspacing='0' cellpadding='1' bordercolor='#800000' bordercolorlight='#800000' bordercolordark='#FFFFFF'><tr align=center><td>场次</td><td>赔率</td><td>数量</td><td>领奖</td></tr>"
do while not (rs.BOF or rs.EOF)
	cname=rs("下注者")
	lj=rs("领奖")
	ci=rs("场次")
	cc=rs("参赛者")
	ss=rs("胜利者")
	sl=rs("数量")
	pl=rs("赔率")
	ab=sl*pl
	if lj<>True and cc=ss then
	msg=msg&"<tr><td >"&rs("场次")&"</td><td>"&rs("赔率")&"</td><td>"&rs("数量")&"</td><td align=center><a href='dui2.asp?mg="&cname&"&bb="&ab&"&changci="&ci&"'>GO</a></td></tr>"
    end if
	rs.MoveNext
loop
msg=msg&"</table>"
rs.Close 
set rs=nothing
connt.Close
set connt=nothing
%>
<head><link rel="stylesheet" href="../../chatroom/css.css">
<title>兑奖</title>
</head>
<body oncontextmenu=self.event.returnValue=false marginwidth="5" marginheight="5" bgcolor="#FFDDF2" background='../bg1.gif'>
<div align=center>
<br>
<b><font  face="宋体"><%=username%>的兑奖表</font></b>
<%=msg%>
<input type=button value='返回' onclick=javascript:location.href='anqi.asp';>
</div>
</body>

