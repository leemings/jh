<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then 
Response.Write "<script language=javascript>{alert('提示：您没有登陆或已经超时断开连接，请重新登陆！');parent.history.go(-1);}</script>" 
Response.End 
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
gwm="千年蝎王"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb") 
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
tl=rs("体力")
nl=rs("内力")
fl=rs("法力")
hydj=rs("会员等级")
dj=rs("等级")
select case hydj
	case 0
		jgsj=320
	case 1
		jgsj=260
	case 2
		jgsj=220
	case 3
		jgsj=180
	case 4
		jgsj=130
	case 5
		jgsj=100
        case 6
		jgsj=60
sj=DateDiff("s",rs("操作时间"),now())
	s=jgsj-sj
end select
if dj<50 and dj>100 then 
	            Response.Write "<script Language=Javascript>alert('提示：您的等级太低。进入初级怪物区的限制是：等级大于50级小于100级。！');location.href = 'javascript:history.go(-1)';</script>"
	            Response.End
end if
if tl<0 then 
	Response.Write "<script Language=Javascript>alert('提示：您已经 挂了。。55555555下次注意点。！');location.href = '../../../exit.asp';</script>"
	Response.End
end if
%>
<HTML>
<HEAD>
<TITLE>梦幻魔界--<%=gwm%></TITLE>
<link href="../../dg/setup.css" rel=stylesheet type="text/css">
<body oncontextmenu=self.event.returnValue=false bgcolor=#000000>
<br><br><br>
<table align=center width="300">
<div align=center><%
if Minute(time())<40 or Minute(time())>=60 then
	Response.Write "<font color=00ffff face=粗体>梦幻魔界尚未开放</font>"
else
	Response.Write "<font color=00ff00 face=粗体>梦幻魔界开始放猎</font>"
end if
%><br><br>
<div align=center><img src=../pic/xz.gif width=160 height=128></div>
<p align=center> <a href=xz1.asp><font color=ffff00><b>发动攻击</b></font></a> <a href=zhx.asp><font color=red><b>唤醒怪物</b></font></a></p>

<p align=left><font color=cccccc size=2>说明：<font color=00ff00>每个小时后二十分钟</font>，可以来梦幻魔界打猎，首先体力内力要大于<font color=ffff00>50000</font>点，注意点别挂就是的，呵呵，打死怪物后会<font color=ffff00>根据怪物的属性级别可以随机获得：锻造武器的物品、积分（存点）、金币、卡片、配药、高级装备<b><b><font color=00ff00>1000</font></b></b>点法力，还要注意您唤醒的怪物其他玩家也可以打，要抓紧时间哦。</font></p>
</table><br><br>
<table align=center width="500">
<div align=center><font color=cccccc><font color=00ffff>猎杀者</font>：<font color=cccccc><%=rs("姓名")%><font> <font color=00ffff>会员等级</font>：<font color=cccccc><%=rs("会员等级")%><font>级 <font color=00ffff>时间限制</font>：<font color=00ff00><b><%=jgsj%></b><font><font color=ffff00> 秒</font> 
<% if s>0 then 
Response.Write "<font color=00ffff>还差</font>：<font color=00ff00><b>"&s&"</b><font><font color=ffff00> 秒</font></font></div>"
else
Response.Write "<font color=00ff00>可以猎杀<font></font></div>"
end if %>
</table><br><br>
<table border="1" width=650 cellspacing="1" cellpadding="0" bordercolor="eeeeee" height="26" align=center>
<tr>
<td height=25 align=center><font color=cccccc>怪物名</font></td>
<td height=25 align=center><font color=cccccc>体力</font></td>
<td height=25 align=center><font color=cccccc>杀气</font></td>
<td height=25 align=center><font color=cccccc>攻击</font></td>
<td height=25 align=center><font color=cccccc>防御</font></td>
<td height=25 align=center><font color=cccccc>等级</font></td>
<td height=25 align=center><font color=cccccc>状态</font></td>
<td height=25 align=center><font color=cccccc>猎杀者</font></td>
</tr>
<% 
rs.close
gwm="千年蝎王"
rs.open "select * from 怪物 where 姓名='"&gwm&"'",conn,2,2
%>
<tr>
<td height=25 align=center><font color=00ff00><%=rs("姓名")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("体力")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("杀气")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("攻击")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("防御")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("等级")%></font></td>
<td height=25 align=center><font color=ffff00><%=rs("状态")%></font></td>
<td height=25 align=center><font color=00ffff><%=rs("猎杀者")%></font></td>
</tr>
<% rs.close %>
</table><br><br><br><br>
<table align=center width="600">
<div align=center><a href=index.asp><font color=ffff00>【返回上页】</font></a> <a href="javascript:location.reload()"><font color=00ff00>【刷新本页】</font></a><br><br><br><font size="2">版权所有 <font color="#FFFFFF">『快乐江湖网』</font></font> 
  <font color="#FF0000" size="2">回首当年</font><font size="2"> 修改制作  请保留版权</font></div>









