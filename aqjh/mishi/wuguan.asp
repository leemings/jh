<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<!--#include file="../config.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("axjh_usermdb")
rs.open "select 体力,内力,武功,武功加,等级,会员等级,会员 from 用户 where 姓名='"&axjh_name&"'",conn
hydj=rs("会员等级")
pdhy=rs("会员")
mytl=rs("体力")
mynl=rs("内力")
mywg=rs("武功")
wgsx=int(rs("等级")*axjh_wgsx+3800+rs("武功加"))
rs.close
set rs=nothing
conn.close
set conn=nothing
select case hydj
	case 0
		bf=1.2
	case 1
		bf=1.25
	case 2
		bf=1.3
	case 3
		bf=1.35
	case 4
		bf=1.4
end select
if pdhy<>false then
	bf=1.3
end if
maxwg=int(wgsx*bf)
wgc=maxwg-mywg
if wgc>0 and wgc=<1000 then
	xlwg="骑马蹲裆"
	xtl=10000
	xnl=5000
elseif wgc>1000 and wgc<=10000 then
	xlwg="少林硬功"
	xtl=100000
	xnl=50000
elseif wgc>10000 and wgc<=100000 then
	xlwg="峨眉心法"
	xtl=1000000
	xnl=500000
elseif wgc>100000 and wgc<=150000 then
	xlwg="逍遥神剑"
	xtl=1500000
	xnl=750000
elseif wgc>150000 and wgc<=200000 then
	xlwg="苗家蛊毒"
	xtl=2000000
	xnl=1000000
elseif wgc>200000 then
	xlwg="浩然正气"
	xtl=10000000
	xnl=5000000
elseif wgc<=0 then
	xlwg="无须修练"
	xtl=0
	xnl=0
end if
if xlwg="无须修练" then
	mess="你目前的武功已达最高境界，无需修练，等升级后再来吧。"
else
	mess="修练“"&xlwg&"”需消耗"
	if xtl-mytl>0 then
		mess=mess&"<font color=red>"&xtl&"点</font>体力，你目前的体力不足。"
	else
		mess=mess&"<font color=red>"&xtl&"点</font>体力，"
	end if
	if xnl-mynl>0 then
		mess=mess&"<font color=red>"&xnl&"点</font>内力，你目前的内力不足。"
	else
		mess=mess&"<font color=red>"&xnl&"点</font>内力。"
	end if
end if
%>
<html>

<head>
<link rel="stylesheet" href="css.css">
<title><%=application("axjh_chatroomname")%></title>
<style type="text/css">
<!--
.style9 {font-size: 12px}
-->
</style>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../jhimg/bgcheetah.gif">
<br>
<br>
<table border="0" cellspacing="0" cellpadding="0" width="366" align="center">
  <tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=axjh_name%></font>欢迎光临习武场</b><br>
        <br>
        以下是你当前的状态<br>
        体力：<font color="#FF0000"><%=mytl%></font>，内力：<font color="#FF0000"><%=mynl%></font><br>
        武功：<font color="#FF0000"><%=mywg%></font>，武功上限：<font color="#FF0000"><%=wgsx%></font></font><br>
        武功最高可修练到：<font color="#008000"><%=maxwg%></font>点。还差<font color="#008040"><%=maxwg-mywg%></font>点。 
        <%if xlwg<>"无须修练" then%>
        <br>
        建议修练：<font color="#FF0000"><%=xlwg%></font> 
        <%end if%>
        <br>
		<%=mess%></div>
<form method=POST action='wuguanok.asp'>
        <table width="414" align="center">
          <tr>
<td>
<tr>
<td align=center>
<select name=money size=1> 
<option value="1000" <%if xlwg="骑马蹲裆" or xlwg<>"无须修练" then%>selected<%end if%>>骑马蹲裆</option>
<option value="10000" <%if xlwg="少林硬功" then%>selected<%end if%>>少林硬功</option>
<option value="100000" <%if xlwg="峨嵋心法" then%>selected<%end if%>>峨眉心法</option>
<option value="150000" <%if xlwg="逍遥神剑" then%>selected<%end if%>>逍遥神剑</option>
<option value="200000" <%if xlwg="苗家蛊毒" then%>selected<%end if%>>苗家蛊毒</option>
<option value="1000000" <%if xlwg="浩然正气" then%>selected<%end if%>>浩然正气</option>
</select>
</td>
</tr>
<tr>
<td align=center><input type=submit value=修习武功 name="submit"></td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
操作简介</div>
</td>
</tr>
<tr>
            <td valign="top" > 
              <p><br>
                <font color="#660000">本江湖密室练武不需要有师傅，所消耗的也不是银两，而是体力和内力。每提升1点武功需消耗<b><font color="#FF0000">10点</font></b>体力和<b><font color="#FF0000">5点</font></b>内力。</font><br>
                (1)、骑马蹲裆 加武功：1000点 &nbsp(2)、少林硬功 加武功：10000点<br> 
                (3)、峨嵋心法 加武功：100000点 (4)、逍遥神剑 加武功：150000点<br> 
                (5)、苗家蛊毒 加武功：200000点 (6)、浩然正气 加武功：1000000点<br> 
                <br>
              <p><%if hydj=0 and pdhy=false then%>
你可以修练武功至上限的<%=bf%>倍！
<%else
if pdhy<>false then%>
你是泡点制会员，可以修练武功超出上限的<%=bf%>倍！
<%else%>
你是<%=hydj%>级会员，可以修练武功超出上限的<%=bf%>倍！
<%end if
end if%></p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<div align="center">
</div>
</body>
</html>
