<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"

if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');</script>"
	Response.End
end if
r=request.querystring("username")
if r<>"" then

Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../liwu/sdlw.mdb")


select case r
	case 1 '清理新人礼物表
		sql="delete * from 新人礼物 where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>【江湖消息】"&aqjh_name&"站长爱护新人,等级等于55级可以领取新人礼物咯" & "</b></font></marquee></marquee>"
	Response.Write "<script Language=Javascript>alert('恭喜,清理成功');</script>"

	case 2 '清理周六礼物
		sql="delete * from 周六礼物 where id>0"
	says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>【江湖消息】"&aqjh_name&"站长清理了周六礼物,这个周六可以领取礼物咯" & "</b></font></marquee></marquee>"
	Response.Write "<script Language=Javascript>alert('恭喜,清理成功');</script>"

	case 3 '清理周末礼物
		sql="delete * from 周日礼物 where id>0"
	says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>【江湖消息】"&aqjh_name&"站长清理了周日礼物,这个礼拜天可以领取礼物咯" & "</b></font></marquee></marquee>"
	Response.Write "<script Language=Javascript>alert('恭喜,清理成功');</script>"

	case 4 '清理月末礼物
		sql="delete * from 月末礼物 where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>【江湖消息】"&aqjh_name&"站长清理了月末礼物,这个月28、29、30号可以领取礼物咯" & "</b></font></marquee></marquee>"
	Response.Write "<script Language=Javascript>alert('恭喜,清理成功');</script>"

	case 5 '清理站长发放的礼物
		sql="delete * from 站长礼物 where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>【江湖消息】"&aqjh_name&"站长给各位侠客准备了丰富的礼物,赶快去领取站长礼物吧" & "</b></font></marquee></marquee>"
			Response.Write "<script Language=Javascript>alert('恭喜,发放成功');</script>"
	
	case 6 '清理周五礼物
		sql="delete * from 周五礼物 where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>【江湖消息】"&aqjh_name&"站长清理了周五礼物,这个周五可以领取礼物咯" & "</b></font></marquee></marquee>"
			Response.Write "<script Language=Javascript>alert('恭喜,清理成功');</script>"
	
	case 7 '清理节日礼物
		sql="delete * from 节日礼物 where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>【发放礼物】教师节到了,"&aqjh_name&"站长给大家准备了丰富的礼物，赶快去领取节日礼物咯!" & "</b></font></marquee></marquee>"
			Response.Write "<script Language=Javascript>alert('恭喜,清理成功');</script>"


case else

	Response.Write "<script Language=Javascript>alert('请从正确的连接进入');</script>"
	Response.End

end select


set rs1=conn1.execute(sql)



set rs1=nothing
conn1.close
set conn1=nothing

call showchat(says)
end if



%>
<html>
<head>
<title>礼物管理</title>
<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body text="#000000" background="../chat/f3.gif">
<div align="center">
<table width="603" border="0" cellspacing="0" cellpadding="0" bordercolor="#0A246A" bgcolor="#808000">
  <tr>

    <td width="5%" align=center ><a href="?username=1"><font color="#800000">清理新人礼物</font></a></td>
    <td width="5%" align=center ><a href="?username=2"><font color="#800000">清理周六礼物</font></a></td>
    <td width="5%" align=center ><a href="?username=3"><font color="#800000">清理周末礼物</font></a></td>
    <td width="5%" align=center ><a href="?username=4"><font color="#800000">清理月末礼物</font></a></td>
    <td width="5%" align=center ><a href="?username=5"><font color="#800000">发放站长礼物</font></a></td>
    <td width="5%" align=center ><a href="?username=6"><font color="#800000">清理周五礼物</font></a></td>
    <td width="5%" align=center ><a href="?username=7"><font color="#800000">发放节日礼物</font></a></td>


  </tr>
</table>
</div>
<p align=center>
请注意:新人礼物是55级可以领取,周五、周六、周日的都是在固定的日期(服务器时间,如果你的服务器时间不准请联系服务器提供商),站长礼物是点了之后就可以让大家去领取的,有的人可能当时没有领取,过了几天才领取,也很正常(所以聊天室有时候会出现某人领了站长的礼物,其实是以前没有领取过)</p>
<p align=center>
　</p>
<p align=center>
此功能新做出来，还没有做自动清理，所以麻烦各位站长定期清理</p>
<p align=center>
注意：新人礼物可以不用清理，其他的请不要在他规定的时间内清理，比如周六那天禁止清理周六礼物，这样会导致重复领取</p>
<p align=center>
站长想发站长礼物的时候就点最后一个，聊天室会有显示的，没有领取的我相信也不是经常上江湖的，所以他们说少领取什么的那是他的事情</p>
</body>
</html>