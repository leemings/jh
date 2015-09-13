<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两,性别,保留,hw from 用户 where 姓名='"&aqjh_name&"'",conn
if rs("性别")="男" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('妇科医院，男人止步！');window.close();}</script>"
	Response.End
end if
hai=rs("保留")
yin=rs("银两")
rs.close
if hai="保留" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你还没有身孕呢，不需要检查！');window.close();}</script>"
	Response.End
end if
if yin<1000000 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('检查一次收取费用100万，你的钱不够！');window.close();}</script>"
	Response.End
end if
hdata=split(hai,"|")
baby=hdata(0)
babysj=hdata(1)
babytl=clng(hdata(2))
babyzt=hdata(3)
erase hdata
sysj=DateDiff("d",babysj,now())
top="今天是你怀孕第"&sysj&"天，你的胎儿一切正常。一定要注意保养身体哟。"
if sysj>=3 and sysj<10 and babyzt<>"药1" then
	babyzt="药"
	top="你今天是怀孕的第"&sysj&"天，你现在需要去服用保胎药了。一次收费500万银两和15金币。"
end if
if sysj>=10 and sysj<20 and baby<>"针1" then
	babyzt="针"
	top="你今天是怀孕的第"&sysj&"天，需要去打保胎针了。一次收费1500万银两和20金币。"
end if
if sysj=20 then
	babyzt="生"
	top="哎呀，你马上就要生了，快去产房吧。快乐医院出于人道主义，对生小孩免收任何费用。"
end if
if sysj=21 then
	babyzt="难"
	top="喂，你都第21天了，超时间了，今天如果不生，那明天你的婴儿就胎死腹中了。快快去医院吧。"
end if
newbl=baby&"|"&babysj&"|"&babytl&"|"&babyzt
if sysj>21 then
	newbl="保留"
	top="哎，你的胎儿在腹中超过21天未生育，已经胎死腹中了。让我们为这尚未出世的小生命的消失来默哀吧……"
end if
conn.execute "update 用户 set 保留='"&newbl&"',银两=银两-1000000 where 姓名='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>怀孕检查</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
  <p>&nbsp;</p>
  <p><font size="7" color="#FF0000">江湖妇科中心-怀孕检查</font></p>
  <p>&nbsp;</p>
  <p><img src="images/zd.gif" width="300" height="144"> <br>
    <font size="5"><%=top%></font></p>
  <p><font size="5"><a href="fkyy.asp">返回</a></font></p>
</div>
</body></html>