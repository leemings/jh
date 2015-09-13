<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<%
sub gfsj(gfdata)
if not(isnull(gfdata)) and instr(gfdata,"|")<>0 then
	mydata=split(gfdata,"|")
	mygj=mygj+mydata(2)
	myfy=myfy+mydata(3)
	erase mydata
end if
end sub
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<!查IP>
<%
ip=request("ip")
if ip="" then
ip=Request.ServerVariables("REMOTE_ADDR")
end if
sip=split(ip,".")
if ubound(sip)<>3 then
ip=Request.ServerVariables("REMOTE_ADDR")
erase sip
sip=split(ip,".")
end if
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
rs.open "select * FROM z WHERE a<=" & num & " and b>=" & num,conn,1,1
if rs.eof or rs.bof then
country="亚洲"
city="未知"
else
country=rs("c")
city=rs("d")
end if
if country="" then country="中国"
if city="" then city="未知"
last="地区:" & country & " 城市:" & city
rs.close
%>
<!查IP>
<%
rs.open "Select * from 用户 where 姓名='" & aqjh_name &"'",conn,3,3
gjsx=rs("等级")*aqjh_gjsx
fysx=rs("等级")*aqjh_fysx
if rs("会员")=true and clng(DateDiff("d",now(),rs("会员结束")))>0 then
	pdstr="<font size=-1>[泡点制会员]"&rs("会员结束")&"</font>"
else
	pdstr=""
end if
wgj=rs("武功加")
nlj=rs("内力加")
tlj=rs("体力加")
mygj=rs("攻击")
myfy=rs("防御")
mywg=rs("武功")
call gfsj(rs("z1"))
call gfsj(rs("z2"))
call gfsj(rs("z3"))
call gfsj(rs("z4"))
call gfsj(rs("z5"))
call gfsj(rs("z6"))
if rs("保留1")<>"保留" and rs("保留1")<>"" then
	zbdata=split(rs("保留1"),"|")
	xs=0
	if UBound(zbdata)=5 then
		sj=datediff("n",zbdata(1),now())
		jgj=zbdata(2)
		jfy=zbdata(3)
		jwg=zbdata(4)
		if sj<=60 then
			mygj=mygj+zbdata(2)
			myfy=myfy+zbdata(3)
			mywg=mywg+zbdata(4)
			xs=1
		end if
	end if
	erase zbdata
end if
%>
<HTML><HEAD><TITLE>个人状态</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="ztimg/css.css" rel=stylesheet type=text/css>
</HEAD>
<BODY background=ztimg/back1.gif bgColor=#000000 leftMargin=0 text=#cccccc topMargin=0 marginheight="0" marginwidth="0" oncontextmenu=window.event.returnValue=false>
<DIV align=center>
<TABLE align=center bgColor=#000000 border=0 cellPadding=6 cellSpacing=1 width=600>
<TBODY>
<TR bgColor=#4f3e39>
<TD align=middle class=title colSpan=7>
<TABLE border=1 borderColor=#000000 cellPadding=0 cellSpacing=0 width="100%">
<TBODY>
<TR>
<TD bgColor=#003333 borderColor=#cccccc>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width="95%">
<TBODY>
<TR>
<TD><FONT color=#ffcc00><%=aqjh_name%></FONT>大侠 ，以下是您的个人状态<IMG align=absMiddle height=16 src="ztimg/icon16.gif" width=16>[<A href="../hcjs/jhzb/index.asp">武器装备</A>] [<a href="../CHANGAN/xiaowu.asp">房屋</a>]
<!查IP>
<form name="form2" method="get" action="">
查IP:<input name="ip" type="text" id="ip" size="15" maxlength="15" style="color: #666666; background-color: #ffffff; text-decoration: blink; border: 1px solid #009900" value="<%=ip%>">
<input type="submit" name="Submit" value="查找" style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; FONT-FAMILY: '宋体'; COLOR: #FFFFFF; FONT-SIZE: 9pt; HEIGHT: 18px">
来自：<%=last%></form>
<!查IP>
</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR>
<TR>
<TD align=middle bgColor=#644e49 rowSpan=3 width="45">
<IMG border=0 height=30 width=30 name=w_r6_c3 src="<%=rs("名单头像")%>">
</TD>
<TD bgColor=#644e49 colSpan=2 align=center>
<FONT color=#ffcc00><%=aqjh_name%></FONT>
</TD>
<TD align=middle bgColor=#51403c width=44>ＩＤ</TD>
<TD align=middle bgColor=#644e49 width="117"><FONT color=#ff0000><%=rs("id")%></FONT></TD>
<TD align=middle bgColor=#51403c width="30">性别</TD>
<TD align=middle bgColor=#644e49 width=141><%=rs("性别")%></TD>
</TR>

<TD align=middle bgColor=#51403c width="30">
<IMG height=1 src="ztimg/spacer.gif" width=30><BR>国家</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("国家")%></TD>
<TD align=middle bgColor=#51403c width="44">
<IMG height=1 src="ztimg/spacer.gif" width=30><BR>职位</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("职位")%></TD>
<TD align=middle bgColor=#51403c width="30">
<IMG height=1 src="ztimg/spacer.gif" width=30><BR>种族</TD>
<TD align=middle bgColor=#644e49 width=141><%=rs("种族")%></TD>
</TR>


<TR>
<TD align=middle bgColor=#51403c width="30">门派</TD>
<TD align=middle bgColor=#644e49 width=117><FONT color=#ffcc00><IMG align=absMiddle height=16 src="ztimg/space.gif" width=18></FONT>
<FONT color=#cccccc><%=rs("门派")%></FONT></TD>
<TD align=middle bgColor=#51403c width="44">身份</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("身份")%></TD>
<TD align=middle bgColor=#51403c width="30"><font color="#cccccc">管理</font></TD>
<TD align=middle bgColor=#644e49 width=141><FONT color=#cccccc><%=aqjh_grade%></FONT></TD>
</TR>
<TR>


<TR bgColor=#433730>
<TD align=middle bgColor=#644e49 rowspan="9" valign="middle"><img src="ztimg/grzt.gif" width="45" height="192"></TD>
<TD align=middle bgColor=#51403c width="30">会员</TD>
<TD align=middle bgColor=#644e49 width="117"><%=rs("会员等级")%></TD>
<TD align=middle bgColor=#51403c width="44">会期</TD>
<TD align=middle bgColor=#644e49 width="117"><%=rs("会员日期")%></TD>
<TD align=middle bgColor=#51403c width="30">年龄</TD>
<TD align=middle bgColor=#644e49 width=141><%=rs("年龄")%></TD>
</TR>
<TR bgColor=#433730>
<TD align=middle bgColor=#51403c width="30">职业</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("职业")%></TD>
<TD align=middle bgColor=#51403c width="44">等级 </TD>
<TD align=middle bgColor=#644e49 width=117><B><%=rs("等级")%></B></TD>
<TD align=middle bgColor=#51403c width="30">练功</TD>
<%if rs("保护")=true then
	bhlg="练功保护"
else
	bhlg="没有保护"
end if%>
<TD align=middle bgColor=#644e49 width=141><%=bhlg%></TD>
</TR>
<TR bgColor=#433730>
<TD align=middle bgColor=#51403c width="30">名号</TD>
<TD align=middle bgColor=#644e49 width=117>
<%
if rs("等级")<20  then response.write("平民")
if rs("等级")>=20 and rs("等级")<30 then response.write("士兵")
if rs("等级")>=30 and rs("等级")<50 then response.write("武士")
if rs("等级")>=50 and rs("等级")<70 then response.write("骑士")
if rs("等级")>=70 and rs("等级")<95 then response.write("斗士")
if rs("等级")>=95 and rs("等级")<120 then response.write("圣武士")
if rs("等级")>=120 and rs("等级")<155 then response.write("圣骑士")
if rs("等级")>=155 and rs("等级")<190 then response.write("圣斗士")
if rs("等级")>=190 and rs("等级")<235 then response.write("究极战士")
if rs("等级")>=235 and rs("等级")<290 then response.write("究极战神")
if rs("等级")>=290 and rs("等级")<345 then response.write("无极战王")
if rs("等级")>=345 and rs("等级")<410 then response.write("王者至尊")
if rs("等级")>=410 and rs("等级")<475 then response.write("天王之王")
if rs("等级")>=475 and rs("等级")<550 then response.write("环宇战仙")
if rs("等级")>=550 then response.write("天外来客")%>
</TD>
<TD align=middle bgColor=#51403c width="44">婚恋</TD>
<TD bgColor=#644e49 colspan="3">结婚：<%=rs("结婚次数")%>次 <%if rs("配偶")<>"无" and rs("配偶")<>"" then%>纪念:<%=rs("结婚记念日")%>
        结婚时间:<%=DateDiff("d",rs("结婚记念日"),date())%>天<%end if%></TD>
</TR>
<TR bgColor=#433730>
<TD align=middle bgColor=#51403c width="30">金币</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("金币")%></TD>
<TD align=middle bgColor=#51403c width="44">金卡</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("会员金卡")%></TD>
<TD align=middle bgColor=#51403c width="30">银两</TD>
<TD align=middle bgColor=#644e49 width=141><%=rs("银两")%></TD>
</TR>
<TR bgColor=#433730>
<TD align=middle bgColor=#51403c width="30">体力</TD>
<TD align=middle bgColor=#644e49 width=117>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width="120" height="13">
<TBODY>
<TR>
<%if rs("体力")<=0 then
	abc=0
elseif rs("体力")>=(rs("等级")*aqjh_tlsx+5260+tlj) then
	abc=120
else
	abc=int(120*(rs("体力")/(rs("等级")*aqjh_tlsx+5260+tlj)))
end if%>
<TD width=120 background="ztimg/000.gif"><img src="ztimg/001.gif" width="<%=abc%>" height="11"></TD>
</TR>
<TR>
<TD width=121 align="center">
<font color=yellow><%=rs("体力")%></font>
<%if aqjh_sx=1 then%>
/<font color=green><%=rs("等级")*aqjh_tlsx+5260+tlj%></font>
<%end if%>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
<TD align=middle bgColor=#51403c width="44">内力</TD>
<TD align=middle bgColor=#644e49 width=117>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="120" height="13">
<TBODY>
<%if rs("内力")<=0 then
	abc=0
elseif rs("内力")>=(rs("等级")*aqjh_nlsx+2000+nlj) then
	abc=120
else
	abc=int(120*(rs("内力")/(rs("等级")*aqjh_nlsx+2000+nlj)))
end if%>
<TR>
<TD width=120 background="ztimg/000.gif"><img src="ztimg/001.gif" width="<%=abc%>" height="11"></TD>
</TR>
<TR>
<TD width=117 align="center">
<font color=yellow><%=rs("内力")%></font>
<%if aqjh_sx=1 then%>
/<font color=green><%=rs("等级")*aqjh_nlsx+2000+nlj%></font>
<%end if%>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
<TD align=middle bgColor=#51403c width="30">武功</TD>
<TD align=middle bgColor=#644e49 width=141>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="120" height="13">
<TBODY>
<%if mywg<=0 then
	abc=0
elseif mywg>=(rs("等级")*aqjh_wgsx+3800+wgj) then
	abc=120
else
	abc=int(120*(mywg/(rs("等级")*aqjh_wgsx+3800+wgj)))
end if%>
<TR>
<TD width=122 background="ztimg/000.gif"><img src="ztimg/001.gif" width="<%=abc%>" height="11"></TD>
</TR>
<TR>
<TD width=122 align="center">
<font color=yellow><%=mywg%></font>
<%if aqjh_sx=1 then%>
/<font color=green><%=rs("等级")*aqjh_wgsx+3800+wgj%></font>
<%end if%>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
</TR>
<TR bgColor=#51403c>
<TD align=middle width="45">攻击</TD>
<TD align=middle bgColor=#644e49><%=mygj%></TD>
<td align=middle>防御</td>
<td align=middle bgColor=#644e49><%=myfy%></td>
<td>存款</td>
<td align=middle bgColor=#644e49><%=rs("存款")%></td>
</TR>
<TR bgColor=#51403c>
<TD align=middle width="45">存点</TD>
<TD align=middle bgColor=#644e49><%=rs("allvalue")%></TD>
<td align=middle>豆点</td>
<td align=middle bgColor=#644e49>
<%if rs("grade")=3 and rs("身份")="护法" then
	response.write (rs("泡豆点数")-int(rs("泡豆点数")/500)*500)&"<font color=#FF0000>/500</font>"
else
	response.write (rs("泡豆点数")-int(rs("泡豆点数")/1000)*1000)&"<font color=#FF0000>/1000</font>"
end if%>
</td>
<td>豆子</td>
<td align=middle bgColor=#644e49>
<%if rs("grade")=3 and rs("身份")="护法" then
	response.write int(rs("泡豆点数")/500)
else
	response.write int(rs("泡豆点数")/1000)
end if%>
</td>
</TR>
<TR bgColor=#51403c>
<TD align=middle width="45">杀人</TD>
<TD align=middle bgColor=#644e49><%=rs("杀人数")%></TD>
<td align=middle>总数</td>
<td align=middle bgColor=#644e49><%=rs("总杀人")%></td>
<td>宝物</td>
<td align=middle bgColor=#644e49><%if rs("宝物")<>"无" and rs("宝物")<>"" then%><font color=red><%=rs("宝物")%></font><%else%><%=rs("宝物")%><%end if%></td>
</TR>
<TR bgColor=#51403c>
<TD align=middle width="45">道德</TD>
<TD align=middle bgColor=#644e49><%=rs("道德")%></TD>
<td align=middle>魅力</td>
<td align=middle bgColor=#644e49><%=rs("魅力")%></td>
<td>师傅</td>
<td align=middle bgColor=#644e49><%=rs("师傅")%></td>
</TR>

<%sj=DateDiff("n",rs("暴豆时间"),now())
if sj<=60 then%>
<TR bgColor=#433730>
<TD bgColor=#51403c colspan="7" height="28" align="center">
<font color="#00FFFF">暴豆威力时间还剩：<%=60-sj%>分</font>
</TD>
</TR>
<%end if%>
<TR bgColor=#000000>
<TD align=middle colspan="7" height="28"><FONT color=#ffffff>&copy; 版权所有 2004-2005 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#ffffff>快乐江湖网</FONT></A></TD>
</TR></TBODY></TABLE></DIV></BODY></HTML>