<!--#include file="jha.asp"-->
<%
name=Session("Ba_jxqy_username")
sql="select * from 用户 where 姓名='" & name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language=vbscript>
  MsgBox "你不是江湖中人"
  location.href = "javascript:history.back()"
</script>
<%
else
	
        if rs("攻击")>80000 or rs("门派")="官府" then
	response.write "你武功已经超过限制,请选择其他动物"
	conn.close
	response.end
end if
%>
<%end if%>
<%
	randomize timer
	r=int(rnd*100)
        if r>65 then
%>
<script language=vbscript>
  MsgBox "没有找到动物"
  location.href = "javascript:history.back()"
</script>
<%end if%>
<html>
<TITLE>新手训练室</TITLE>
<head>
<style>
input, body, select, td{font-size:14;line-height:160%}
</style>
</head>

<body oncontextmenu=self.event.returnValue=false>
<p align=center style='font-size:16;color:yellow'>欢迎<%=session("myname")%>来此训练
<table border=1 bgcolor="#FFCC99" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#BEE0FC>
<table>
<tr><td>
<form method=POST action='dog1ok.asp'>
<tr><td align=center>选择内力攻击：
<select name=nl size=1> 
<option value="0">0点内力
<option value="10">10点内力
<option value="20">20点内力
<option value="30">30点内力
<option value="40">40点内力
<option value="50">50点内力
<option value="60">60点内力
<option value="70">70点内力
<option value="80">80点内力
<option value="90">90点内力
<option value="100">100点内力
<option value="110">110点内力
<option value="120">120点内力
<option value="130">130点内力
<option value="140">140点内力
<option value="150">150点内力
<option value="200">200点内力
<option value="300">300点内力
<option value="500">500点内力
<option value="700">700点内力
</select></td></tr>
<tr><td colspan=2 align=center><input type=submit value=确定>
<input type=button value=返回 onclick="location.href='xuetang.asp'">
</td></tr>
<tr><td colspan=2 style='font-size:9pt'><hr>
  狼狗 不但会主动攻击， 还比较会闪躲，敏捷， 若自己的活力消耗到一定的程度就会逃跑， 建议初期玩家可打狗修练。防御=0-35点</td></tr>
</form>
</table>
</table>

</body>
</html>