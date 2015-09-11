<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=210"
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
if regjm<>regjm1 then
	%>
	<script language=vbscript>
	alert ("你输入的认证码不正确，应该输入:<%=regjm%>")
	location.href = "javascript:history.back()"
	</script>
	<%
	response.end
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	%>
	<script language="vbscript">
	 alert "你不能进行操作，进行此操作必须进入聊天室！"
	 window.close()
	</script>
	<%
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,身份,密码 from 用户 where 姓名='"&aqjh_name&"'",conn
if rs("身份")<>"掌门" then
	%>
	<script language="vbscript">
	 alert "你又不是掌门，乱跑什么！"
	 window.close()
	</script>
	<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
pass=request.form("password")
pass=trim(pass)
to1=trim(request.form("xrzm"))
if instr(to1,"or")<>0 or instr(pass,"or")<>0 then Response.Redirect "../error.asp?id=54"
if to1="无" or to1="无" or to1="未定" or to1="未定" then Response.Redirect "../error.asp?id=130"
if left(to1,3)="%20" OR InStr(to1,"=")<>0 or InStr(to1,"`")<>0 or InStr(to1,"'")<>0 or InStr(to1," ")<>0 or InStr(to1,"　")<>0 or InStr(to1,"'")<>0 or InStr(to1,chr(34))<>0 or InStr(to1,"\")<>0 or InStr(to1,",")<>0 or InStr(to1,"<")<>0 or InStr(to1,">")<>0 then Response.Redirect "../error.asp?id=120"
if to1=Application("aqjh_automanname") or to1="大家" or Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能传位给机器人或大家，或者你要传位的人不在聊天室里？');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if pass="" then	Response.Redirect "../../error.asp?id=432"
if aqjh_name=to1 then
	Response.Write "<script Language=Javascript>alert('你已经是掌门了，还给自己传什么位呀，是不是喝醉了？');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
%><!--#include file="../../pass.asp"--><%
pass=md5(pass)
if rs("密码")<>pass then
	%>
	<script language="vbscript">
	 alert "对不起，你的密码错误！"
	 window.close()
	</script>
	<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
if rs("门派")="官府" then
	Response.Write "<script Language=Javascript>alert('你要做什么？！');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
pai=rs("门派")
rs.close
rs.open "SELECT 门派,等级,身份,grade FROM 用户 WHERE trim(姓名)='" & to1 & "'",conn
if rs.bof or rs.eof then
	Response.Write "<script Language=Javascript>alert('你要传位给谁呀，江湖中好象没有这个人！');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
if rs("grade")>=6 or rs("门派")="官府" then
	Response.Write "<script Language=Javascript>alert('他是官府人员或已经是掌门,不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
if rs("门派")<>pai or rs("等级")<35 then
	Response.Write "<script Language=Javascript>alert('他的等级不足35级或他不是你门下弟子，不能禅位！');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 等级>=18 and 介绍人='"& to1 &"'")
lr=tmprs("数量")
set tmprs=nothing
if lr<3 then
	Response.Write "<script Language=Javascript>alert('提示：他的拉人记录不足3人，或你所拉的人的等级还没有到18级以上！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "select 会员等级 from 用户 where 姓名='" & to1 &"'",conn,2,2
hy=rs("会员等级")
if hy=rs("会员等级")and (hy=1 or hy=2 or hy=3 or hy=4 or hy=5) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('对方是等级会员不能禅位给他,有什么问题请找站长说去!');window.close();}</script>"
		response.end
end if
mess="<font color=red>"&pai&"掌门"&aqjh_name&"现将掌门之位传于其门下弟子"&to1&"!</font></marquee>"
conn.execute "update 用户 set 身份='弟子',grade=1,领钱=now() where 姓名='" & aqjh_name & "'"
conn.execute "update 用户 set 身份='掌门',grade=5,领钱=now() where 姓名='" & to1 & "'"
conn.execute "update p set b='"&to1&"' where a='"&pai&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【小道消息】</font><b><font color=green>"&mess&"</font></b>"			'聊天数据
	says=replace(says,"'","\'")
	says=replace(says,chr(34),"\"&chr(34))
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
Response.Write "<script Language=Javascript>alert('提示：恭喜！你掌门禅位成功，请退出重新进入江湖！');window.close();</script>"
response.end
%>