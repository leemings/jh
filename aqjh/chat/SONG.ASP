<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/czdj.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
if aqjh_name="" or Session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj1=n & "-" & y & "-" & r
sj2=s & ":" & f & ":" & m
sj3=sj1 & " " & sj2
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
Set Rs=conn.Execute("select  * FROM song")
if chatbgcolor="" then chatbgcolor="008888"%>
<html>
<head>
<title>点歌</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<script language="javaScript">
function send(){sendurl="songplay.asp?name=" + document.forms[0].song.options[document.forms[0].song.selectedIndex].value;window.location.href=sendurl;}
</script>
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" text="#FFFFFF">
<div align=center><font color="#FFFF00" style="font-size:9pt">点歌系统 | <a href=mtv/swflist.asp><font color="#FFFF00" style="font-size:9pt">MTV点播</font></a></font></div>
<hr size=1 color=FFFF00>
<table border="0" width="100%">
<form method="post" action="songplay.asp?" name="" target="ps">
<tr>
<td>
<p> <font color="#FFFF00" style="font-size:9pt">选择MP3列表:</font> <%if aqjh_grade>=9 then%><a href="#" onClick="window.open('songadd.asp','addmp3','scrollbars=yes,resizable=no,width=450,height=500')"><font color="#FFFF00" style="font-size:9pt">管理</font></a><%end if%><br>
<select name="song" style="font-size:9pt" size="1">
<%do while not rs.bof and not rs.eof %>    
<option value="<%=rs("name")%>"><%=rs("name")%></option>
<%
rs.movenext
loop
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
erase mid    
</select>
</p>
<p><font color="#FFFF00" style="font-size:9pt">播放方式：</font><br>
<input type="radio" name="loopok" value="1" checked>
<font style="font-size:9pt">只听一遍<br>    
<input type="radio" name="loopok" value="infinite">
<font style="font-size:9pt">百听不厌<br>    
<input type="radio" name="loopok" value="ddj">
<font style="font-size:9pt">点给大家<br>    
&nbsp;&nbsp;&nbsp; 需要<%=xyyl%>个金币<br>    
<input type="radio" name="loopok" value="dhy">
<font style="font-size:9pt">点给好友<br>    
&nbsp;&nbsp;&nbsp; 需要<%=pyyl%>个金币<br>    
<br>
祝福话语<br>
<!--webbot bot="Validation" B-Value-Required="TRUE" I-Maximum-Length="50" --><input type="text" name="zhufu" size="17" value="祝大家每天都有一个好心情!" maxlength="50"><br>
朋友名字<br>
    <select name="to1">
      <%useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
for i=0 to x
if show(i)<>aqjh_name Then
%>
      <option value="<%=show(i)%>" selected><%=show(i)%></option>
<%end if
next%>
<%if x=0 then%>
      <option value="就你自己一个人" selected>就你自己一个人</option>
<%end if%>
    </select>


 　</p>
<table border="0" cellpadding="4">
<tr>
<td>
<input type="submit" name="play" value="播放">
</td>
<td align="right">
<input type="submit" name="st" value="停止">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
<Script Language=Javascript>
parent.m.location.href='about:blank';
</Script>
</body>
</html>
