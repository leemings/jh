<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
id=trim(request("id"))
if instr(id,"官")<>0 then
		Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');window.close();}</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 等级,魅力,道德,银两,门派 from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then	
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：江湖上找不到你的资料！');window.close();</script>"
		response.end
end if
if rs("等级")<65 or rs("魅力")<100000 or rs("道德")<100000 or rs("银两")<1000000000 then
	Response.Write "<script Language=Javascript>alert('提示：加入官府一定要战斗等级超过65级,魅力和道德超过10万,银两10亿！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("门派")="官府" then
	Response.Write "<script Language=Javascript>alert('提示：你现在已经是官府了！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 等级>=20 and 介绍人='"& sjjh_name &"'")
lr=tmprs("数量")
set tmprs=nothing
if lr<20 then
	Response.Write "<script Language=Javascript>alert('提示：你的拉人记录不足20人，或你所拉的人的等级还没有到20级以上！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 会员等级 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,3,3
hy=rs("会员等级")
if hy=rs("会员等级") then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('等级会员无权加入官府,有什么问题请找站长说去!');window.close();}</script>"
		response.end
end if
	message="你成功地加入了官府这个门派!"
	conn.execute "update 用户 set grade=6,门派='官府',身份='弟子' where 姓名='"&sjjh_name&"'"
says="<font color=#ff0000>【加入官府】" & sjjh_name & "成功的加入了官府这个门派！</font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>〖<%=Application("sjjh_chatroomname")%>〗</title>
<style type="text/css">td           { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt }
select       { font-family: 宋体; font-size: 9pt }
a            { color: #0000FF; font-family: "宋体"; font-size: 9pt; text-decoration: none }
a:hover      { color: #0000FF; font-family: "宋体"; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font color="#FF0000" size="+3">加入官府门派</font></p>
<p align="center"><font color="#FFFFFF"><b><font color="#000000"><%=Application("sjjh_chatroomname")%>告示：<%=message%>
</font></b></font><font color="#000000"></font></p>
<hr>
</body>
</html>
