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
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="data.asp"-->
<!--#include file="func.asp"-->
<%
sh=request.form("sh")
'sy=request.form("sy")
my=aqjh_name
if request.form("h")="1" then
sql="select * from 用户 where 姓名='" & aqjh_name & "'"
set rs=conn.execute(sql)
if rs("体力")<-1000 or rs("状态")="死" then
	Response.Redirect "../../chat/chaterr.asp?id=001"
	rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	response.end
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('警告：请等"& s &"秒再冒险,你想没命呀！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("体力")<20 then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你疲劳程度已超出范围，为防不测，还是离开孤岛为上！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sql="update 用户 set 体力=体力-20 where 姓名='" & aqjh_name & "'"
	conn.execute sql
        rs.close
        set rs=nothing
	conn.close
        set conn=nothing
        message=huayuan(my,sh,sy)
end if
says="<font color=#ff0000>【孤岛冒险】</font>"&message
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
Response.Write "<script language=JavaScript>{alert('提示："&message&"');location.href='index.asp'}</script>"
	Response.End 
%>