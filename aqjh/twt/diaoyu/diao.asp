<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('刷钱是吗？喜欢是吗？点啊，刷啊！！');i=i+1;}top.location.href='../../exit.asp'}</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<8 then
ss=8-sj
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来钓鱼吧！');window.close();</script>"
	Response.End
end if
if Hour(time())<>9 and Hour(time())<>14 then
	Response.Write "<script Language=Javascript>alert('提示：钓鱼的最好时段是上午的9点，下午3点！');window.close();</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>钓鱼</title>
<meta http-equiv="refresh" content="10">
<style></style>
</head>
<BODY oncontextmenu=self.event.returnValue=false BGCOLOR="#fefefe" text="#FFFFFF">
<DIV ID="Layer1" STYLE="position:absolute; left:46px; top:139px; width:192px; height:165px; z-index:3">
<%
if Session("diaoyu")=true then
	Session("diaoyu")=now()
end if 
abc=DateDiff("s",Session("diaoyu"),now())
if abc>=115 then
%>
<div align="center">
<script language=vbscript>
location.href = "pao.asp"
</script>
<%
response.end
end if
if abc<=90 then
%>
<IMG SRC="diaoyu.gif" width="280" height="159"><br>
<font color="#0000FF">请等吧，鱼会有的！</font>
<%
else
%>
<a href="diaoyuok.ASP"><IMG SRC="diaoyuok.gif" border="0" width="280" height="159"></a>
<font color="#0000FF">鱼咬钩了，快起竿！</font>
<%
end if
%>
</div>
</DIV>
<DIV ID="Layer2" STYLE="position:absolute; left:11px; top:13px; width:252px; height:102px; z-index:1"><IMG SRC="diao1.jpg"></DIV>
<div id="Layer4" style="position:absolute; left:263px; top:13px; width:105px; height:66px; z-index:4"><img src="diao2.gif"></div>
<span id="liveclock" style="position: absolute; left: 328px; top: 275px; width: 76px; height: 23px">
<div align="center"><font color=red>进行时间:<br>
<%=abc%>秒</font> </div>
</span></body></html>