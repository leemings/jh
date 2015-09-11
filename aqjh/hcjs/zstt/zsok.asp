<%@ LANGUAGE=VBScript%>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
jifen=rs("allvalue")
myzs=rs("转生")
ztdj=200+myzs*20
ztjf=ztdj^2*50
if jifen<ztjf or aqjh_jhdj<ztdj then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：你现在准备"&myzs+1&"转,需要"&ztdj&"级才行');window.close();}</script>"
		response.end
end if
neili=int(jifen*.5)
tili=int(jifen*.5)
wg=int(jifen*.25)
conn.execute "update 用户 set 内力加=内力加+"&neili&",体力加=体力加+"&tili&",武功加=武功加+"&wg&",allvalue=allvalue-"&ztjf&",等级=等级-"&ztdj&",转生=转生+1 where 姓名='"&aqjh_name&"'"
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','阎王','人命',now(),'转世投胎')"
zstt=rs("转生")
rs.close
set rs=nothing	
set conn=nothing
says="<font color=green>【转世投胎】<font color=#ff0000>" & aqjh_name & "</b>大侠转世投胎了，内力加、体力加各加"&neili&"点，武功加上升"&wg&"点！现在是第"&myzs+1&"次转生</font>"			'聊天数据
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
call showchat(says)
Response.Write "<script language=JavaScript>{alert('提示：您转世投胎好了，点击确定重新登陆江湖！');top.location.href='../../exit.asp';}</script>"
response.end
%>