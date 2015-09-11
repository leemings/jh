<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../../bg.gif">
<%my=aqjh_name
money=abs(request.form("money"))
if  money<>10000000  then 
	Response.Write "<script Language=Javascript>alert('你想作什么？！');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	coon.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=210"
	response.end
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<1 then
	s=1-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"分钟]再操作！');}</script>"
	Response.End
end if	
if rs("银两")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你没有那么多钱呀，你想作什么？！');window.close();</script>"
	response.end
end if
if rs("粪库")=0 then
  rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你钱多吗?粪库为空,还清什么清呀！');window.close();</script>"
	response.end
end if
conn.execute "update 用户 set 粪库=0,银两=银两-"& money &",操作时间=now() where 姓名='"&aqjh_name&"'"
says="<font color=red>【清理粪库】["&aqjh_name&"]</font><font color=green>向江湖收费站交了"&money&"银两,粪库清理完毕!</font>"
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
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
Response.Write "<script Language=Javascript>alert('您向江湖收费站交了"& money &"两，粪库清理完毕！');window.close();</script>"
rs.close
conn.close
set rs=nothing
set conn=nothing
%>