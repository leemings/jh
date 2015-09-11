<%@ LANGUAGE=VBScript%>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,任务,银两,体力,法力,知质,会员金卡,仙术,w6,金,木,水,火,土,操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<10 then
	ss=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来提交任务！');window.close();</script>"
	Response.End
end if
if aqjh_jhdj<120 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('江湖等级需要超过120级才可提交任务！');window.close();</script>"
	Response.End
end if
input=request.form("input")
if rs("木")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你没有1000木属性无法完成任务！');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("金")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你没有1000金属性无法完成任务！');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("水")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你没有1000水属性无法完成任务！');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("仙术")<5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你是神仙也！怎么连仙术都没有，就5点，快去修炼吧！');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("火")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你没有1000火属性无法完成任务！');location.href = 'rwbz.asp';}</script>"
	response.end
end if
input=request.form("input")
if rs("土")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 	Response.Write "<script language=JavaScript>{alert('你没有1000土属性无法完成任务！');location.href = 'rwbz.asp';}</script>"
	response.end
end if
if trim(rs("任务"))<>"神族任务" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的任务不是神族任务，所以您不能进行神族任务操作！！');window.close();</script>"
	response.end
end if
if rs("银两")<1000000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('请交1000000两任务提交费！');window.close();}</script>"
Response.End
end if
if rs("体力")<100000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('体力需要100000！这年头做什么都需要出力的..');window.close();}</script>"
Response.End
end if
if rs("法力")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('需要1000法力才可以完成任务！');window.close();}</script>"
Response.End
end if
if rs("知质")<200 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('看你聪明不？上缴200知质！');window.close();}</script>"
Response.End
end if
conn.execute "update 用户 set 银两=银两-10000 where 姓名='"& aqjh_name &"'"
conn.execute "update 用户 set 体力=体力-1000 where 姓名='"& aqjh_name &"'"
conn.execute "update 用户 set 法力=法力-1000 where 姓名='"& aqjh_name &"'"
conn.execute "update 用户 set 知质=知质-200 where 姓名='"& aqjh_name &"'"
conn.execute "update 用户 set 木=木-1000 where 姓名='"& aqjh_name &"'"
conn.execute "update 用户 set 金=金-1000 where 姓名='"& aqjh_name &"'"
conn.execute "update 用户 set 水=水-1000 where 姓名='"& aqjh_name &"'"
conn.execute "update 用户 set 火=火-1000 where 姓名='"& aqjh_name &"'"
conn.execute "update 用户 set 土=土-1000 where 姓名='"& aqjh_name &"'"
conn.execute "update 用户 set 仙术=仙术-5 where 姓名='"& aqjh_name &"'"
randomize timer
r=int(rnd*9)+1
jinbi=0
jinka=0
zstemp=rs("w6")
select case r
case 1
	dim js(10)
	js(0) ="骨头"
	js(1) ="蓝宝石"
	js(2) ="红宝石"
	js(3) ="僵尸血"
	js(4) ="僵尸牙"
	js(5) ="水晶石"
	js(6) ="绿宝石"
	js(7) ="硫磺"
	js(8) ="硝酸"
	js(9)="木炭"
	randomize()
	myxy = Int(Rnd*10)
	zstemp=add(rs("w6"),js(myxy),5)
	mess="衰噢!!!!!"& aqjh_name &"完成任务得到[<font color=blue>"&js(myxy)&"</font>]5块，可作为配药系统必备的好东西 ！"
	sql="update 用户 set 操作时间=now(),w6='"&zstemp&"' where 姓名='"&aqjh_name&"'"
	erase js
case 2
	mess=""& aqjh_name &"完成任务道德提高1200，攻击提高500"
	sql="update 用户 set 操作时间=now(),道德=道德+1200,攻击=攻击+500 where 姓名='"&aqjh_name&"'"
case 3
	mess=""& aqjh_name &"完成任务道德提高1000"
	sql="update 用户 set 操作时间=now(),道德=道德+1000 where 姓名='"&aqjh_name&"'"
case 4
	mess=""& aqjh_name &"完成任务道德+1300"
	sql="update 用户 set 操作时间=now(),道德=道德+1300 where 姓名='"&aqjh_name&"'"
case 5
	mess=""& aqjh_name &"完成任务道德+1800"
	sql="update 用户 set 操作时间=now(),道德=道德+1800 where 姓名='"&aqjh_name&"'"
case 6
	mess=""& aqjh_name &"完成任务防御+120"
	sql="update 用户 set 操作时间=now(),防御=防御+120 where 姓名='"&aqjh_name&"'"
	case 7
	mess=""& aqjh_name &"完成任务防御+150"
	sql="update 用户 set 操作时间=now(),防御=防御+150 where 姓名='"&aqjh_name&"'"
	case 8
	mess=""& aqjh_name &"完成任务防御+180"
	sql="update 用户 set 操作时间=now(),防御=防御+180 where 姓名='"&aqjh_name&"'"
        case 9
        rs.close
	rs.open "select 姓名 from 用户 where 姓名='"&aqjh_name&"' and 宝物='" & Application("aqjh_baowuname") & "'",conn
	if rs.eof or rs.bof  then
	    mess=""& aqjh_name &"运气真是好，完成任务得到江湖至宝"&Application("aqjh_baowuname")&"！大家还不快抢！"
            sql="update 用户 set 保护=false,宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where 姓名='"&aqjh_name &"'"
        else
            mess=""& aqjh_name &"忽然眼前一亮，居然看到前面不远出现江湖至宝"&Application("aqjh_baowuname")&"！"
            sql="update 用户 set 宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where 姓名='"&aqjh_name &"'"
	end if
end select
conn.execute ""&sql&""
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>【完成任务】</font>"&mess			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="#0000FF"
saycolor="#0000FF"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="jh.css" rel=stylesheet type="text/css">
<title>完成任务</title></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#000000">
<div align="center"> <br>
<br>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr>
<td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr>
<td valign="top">
<div align="center">
<p><%=mess%></p>
</div>
</table>
<div align="center"><br>
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</div>
</td>
</tr>
</table>
</div>
</body></html>