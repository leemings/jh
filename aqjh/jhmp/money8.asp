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
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
n=Year(date())
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r
if day(date())=28 then
	Response.Write "<script Language=Javascript>alert('今天是领月薪水的时间，你怎么还来这里啊！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where allvalue>100 and 介绍人='"&aqjh_name&"'")
jsren=tmprs("数量")
rs.open "Select * from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
if rs("等级")>=20 then
   rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('你的等级已经大于20级了,还来这里干嘛?');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<15 and aqjh_grade<10 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('为防止作弊，刚上线15分钟的人领取不了月薪水!');location.href = 'javascript:history.back()';}</script>"
		response.end

end if
if rs("会员等级")<>0 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('等级会员请不要到这里来!');location.href = 'javascript:history.back()';}</script>"
response.end
end if
if rs("会员")=true and clng(DateDiff("d",now(),rs("会员结束")))>0  then
   rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('泡分制会员请不要到这里来！');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
if Session("aqjh_grade")>=5 then
 Response.Write "<script language=JavaScript>{alert('这不是掌门和官府应该来的地方！');location.href = 'javascript:history.back()';}</script>"
 Response.End
end if
mdate=rs("领钱")
'if CDate(mdate)=now() then
'if day(mdate)>=day(now()) and month(mdate)=month(now()) and year(mdate)=year(now()) then 
if  DateDiff("d",rs("领钱"),now())=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('天呀，["&aqjh_name &"]今天你来领过钱的！忘记了？');window.close();</script>"
	response.end
end if
yin=rs("银两")+rs("存款")
if yin>(rs("会员等级")+1)*500000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('天呀，["&aqjh_name &"]你现在有这么多的银两你还想要呀！！');window.close();</script>"
	response.end
end if
dj=rs("等级")
hy=rs("会员等级")
if trim(rs("身份"))="掌门" then
	'如是掌门统计人数
	tmprs=conn.execute("Select count(*) As 数量 from 用户 where 门派='"&rs("门派")&"'")
	renshu=tmprs("数量")
	'算钱数
	money=cint(renshu)*100000+jsren*50000+dj*50000+500000
	conn.execute("update 用户 set 银两=银两+"&money&",领钱=now() where 姓名='"&aqjh_name&"'")
else
	'个人按会员乘10万
	money=jsren*50000+dj*50000+500000+500000
	conn.execute("update 用户 set 银两=银两+"&money&",领钱=now() where 姓名='"&aqjh_name&"'")
end if
rs.close
rs.open "Select * from 用户 where 姓名='"&aqjh_name&"'",conn
%>
<html>
<head>
<title>工资领取处</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
TD {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
BODY {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
SELECT {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
A {COLOR: #FFC106; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</STYLE>
</head>
<body bgcolor="#CCCCCC" text="#000000" leftmargin="0" background="../jhimg/bk_hc3w.gif">
<div align="center">
<p><font color="#000000"><b><font size="+3">工资领取处</font></b></font> <br>
<br>
<font size=+1> <b><%=aqjh_name%></b></font>你是<font color="#0000FF">[<%=rs("门派")%>]</font>
的<font color="#FF0000"> [<%=rs("身份")%>] </font>拉人：<font color="#FF0000"><b><%=jsren%>个</b></font>    
战斗等级:<font color="#FF0000"><%=rs("等级")%>级</font>    
<%if trim(rs("身份"))="掌门" then%>    
你有弟子:<font color="#FF0000"><b><%=renshu%>个</b></font> <%end if%><br>多多努力，多多拉人！    
<br>
今天你从江湖领取到了<b><font color="#FF0000"><%=money%>两</font></b>，小心保存，不要乱花！    
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
says="<font color=red>【领取工资】["&aqjh_name&"]</font><font color=blue>今天刚领了<font color=red>"&money&"</font>两的工资,大家还不快去,等一下就没有了!</font>"
call showchat(says)
%>
<br>
</p>
<p align="center"><br>
<font color="#FF0000" size="+1"><b>工资计算方法</b></font><br>
<font color="#0000FF">弟子：</font>介绍人数x5万+战斗等级x5万+50万</p>
<p align="center">&nbsp;</p>
<p align="center">
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</p>
</div>
</body>
</html>