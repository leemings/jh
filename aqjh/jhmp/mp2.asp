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
id=trim(request("id"))
if instr(id,"官")<>0 then
		Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');window.close();}</script>"
		Response.End
end if
my=trim(request("my"))
if my<>aqjh_name then
		Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');window.close();}</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 身份<>'掌门' and 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：江湖上找不到你的资料或者你是掌门！');window.close();</script>"
		response.end
end if
	if rs("门派")="" or rs("门派")="游侠" or rs("门派")="无" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：你并无门派！！');window.close();</script>"
		response.end
	end if
	if rs("门派")="出家" or rs("门派")="天网"  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：你是出家人，或是天网杀手不可以在这里离开!！！');window.close();</script>"
		response.end
	end if

if rs("会员等级")=0 then
	message="你离开了" & rs("门派") & "，损失内力、武功、体力各1000,银两及存款均降到1000两！"
	if rs("存款")>1000 then
		conn.execute "update 用户 set 门派基金=0,离派=离派+1,门派='游侠',身份='弟子', 内力=内力-1000,体力=体力-1000,武功=武功-1000,grade=1,银两=1000,存款=1000 where 姓名='"&aqjh_name&"'"
	else
		conn.execute "update 用户 set 门派基金=0,离派=离派+1,门派='游侠',身份='弟子', 内力=内力-1000,体力=体力-1000,武功=武功-1000,grade=1,银两=0 where 姓名='"&aqjh_name&"'"
	end if
else
	message="你离开了" & rs("门派") & "，因为你是江湖会员，所以原有内力，金钱不变！"	
	conn.execute "update 用户 set 门派基金=0,离派=离派+1,门派='游侠',身份='弟子',grade=1 where 姓名='"&aqjh_name&"'"
end if
	conn.execute "update p set c=c-1 where a='" & id & "'"
says="<font color=#ff0000>【判离帮派】" & aqjh_name & "成功的离开了["& id &"]这个门派，银两存款减少到1000，离派次数加1！门派贡献减为0！</font>"			'聊天数据
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
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>〖<%=Application("aqjh_chatroomname")%>〗</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body background="../jhimg/bk_hc3w.gif" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font size="+3" color="#FF0000">离开门派</font><br>
<br>
<br>
<%=Application("aqjh_chatroomname")%>告示：<%=message%> </div>
<hr>
<div align="center">版权所有『快乐江湖总站』</div>
</body>
</html>
