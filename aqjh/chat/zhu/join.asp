<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="../../showchat.asp" -->
<%
aqjh_name=Session("aqjh_name")
username1=request("name1")
username2=request("name2")
if aqjh_name<>username1 then
if aqjh_name=username2 then
response.Write "<script language=JavaScript>alert('提示：您自己邀请别人加入的组，自己已经是组长了！');window.close();</script>"
response.End()
else
response.Write "<script language=JavaScript>alert('提示：您看看清楚，人家是要请你加入的吗？');window.close();</script>"
response.End()
end if
end if
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示:你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
act=1
nowinroom=session("nowinroom")
addwordcolor="660099"
saycolor="008888"
towho="大家"
addsays=aqjh_name

ffsj=application("aqjh_jointime")
nowsj=DateDiff("s",ffsj,now())
if nowsj>=30 then
        Application.Lock
	Application("aqjh_jointime")=""
        Application.UnLock
	Response.Write "<script language=JavaScript>{alert('【快乐江湖网】提示：邀请加入组有效时间30秒，过期作废。');window.close();}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs1.open "select * from 用户 where 姓名='"&username2&"'",conn,1,3
zid=rs1("id")
zname=rs1("组名")
rs1.close
rs.open "select * from 用户 where 姓名='"&username1&"'",conn,1,3
zdzt=rs("组队")
zdm=rs("组名")
if zdzt=false then
response.Write "<script language=JavaScript>alert('提示：您的组队状态还没打开，怎么能加入到组呢？');window.close();</script>"
response.End()
rs.close
conn.close
set rs=nothing
set conn=nothing
end if
if zdm<>"" and zdm<>"无" then
response.Write "<script language=JavaScript>alert('提示：您已经加入组了，不能加入其他组！');window.close();</script>"
response.End()
rs.close
conn.close
set rs=nothing
set conn=nothing
end if
rs("组名")=zname
rs.update
rs.close
rs.open "select * from 组队 where 组长="&zid&" and 组名='"&zname&"'",conn,1,3
if rs.eof and rs.bof then
response.Write "<script language=JavaScript>alert('提示：不存在的组!');window.close();</script>"
response.End()
rs.close
conn.close
set rs=nothing
set conn=nothing
end if 
zttname=rs("组员")
rs("组员")=zttname&","&username1
rs("组员数")=rs("组员数")+1
rs("时间")=now()
rs.update
rs.close
conn.close
set rs=nothing
set conn=nothing
diaox="<font color=#ff0000>【组队消息】</font> <font color=#ff00ff>" & username1 & "</font>应组长<font color=#ff00ff>"&username2&"</font>的邀请,加入了组["&zname&"]!"
call showchat(diaox)
Response.Write "<script language=JavaScript>{alert('加入组成功！');location.href = window.close();}</script>"
response.End()
%>
