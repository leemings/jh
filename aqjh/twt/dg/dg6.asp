<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 性别,体力,操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=454"
	response.end
end if
if rs("性别")<>"女" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你个男人来这里作什么不想活了！');location.href = 'javascript:history.back()';}</script>"
	Response.End 
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<30 then
ss=30-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('刚刚打过工，还在休息中，请等："&ss&"分后再来！');location.href = 'javascript:history.back()';}</script>"
	Response.End 
end if
if rs("体力")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的体力不够500，不可以在这里工作！');location.href = 'javascript:history.back()';}</script>"
	Response.End 
end if
conn.execute "update 用户 set 体力=体力-500,银两=银两+10000,操作时间=now() where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【怡红院消息】</font><font color=green>["&aqjh_name&"]今天在怡红院里做了一天的活，获得银两10000！体力下降500</font>"
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
<title>快乐江湖</title></head>
<body bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr>
          <td height="37"> <font color="#000000"><strong>快乐江湖打工:</strong></font> 
        <tr>
<td height="182" valign="top">在怡红院作了一天的活，你获得银两4000！体力下降100！<Br><Br>
          </td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="返 回" onclick="location.href='dg.asp'">
</div></td></tr></table></td></tr></table></body></html>