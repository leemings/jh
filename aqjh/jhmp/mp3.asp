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
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
you=trim(request("you"))
if instr(you,"OR")<>0 or instr(you,"Or")<>0 or instr(you,"oR")<>0 or instr(you,"or")<>0 or instr(you,">")<>0 or instr(you,"<")<>0 or instr(you,"=")<>0 or instr(you,chr(13))<>0 or instr(you,"'")<>0 or instr(you," ")<>0 or instr(you,"%20")<>0 or instr(you,chr(13))<>0 then
	Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');window.close();}</script>"
	Response.End
end if
if you=aqjh_name then
	Response.Write "<script language=javascript>alert('还有自已赶自已出门派的？？');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 门派 FROM 用户 WHERE 身份='掌门' and 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('提示：你不是这个门派掌门！！');window.close();</script>"
	response.end
end if
rs.close
rs.open "select 门派,身份 from 用户 where 姓名='" & aqjh_name &"'",conn
if rs("门派")="官府" and rs("身份")="掌门" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：官府和掌门不是你要踢出门派的对象！！！');}</script>"
	Response.End
end if
mp=rs("门派")
rs.close
rs.open "SELECT * FROM 用户 WHERE 姓名='"&you&"' and 门派='"&mp&"'" ,conn
if rs.eof or rs.bof then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('提示：["&you&"]并不是你门派弟子！！！');window.close();</script>"
	response.end
end if
message="成功的把" & you & "逐出" & id & "，这样的弟子还是不收的好！"
conn.execute "update p set c=c-1 where a='" & mp & "'"
conn.execute "update 用户 set 门派基金=0,门派='游侠',身份='弟子',grade=1 where 姓名='" & you & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>【逐出弟子】" & aqjh_name & "将自己弟子："& you &"赶出了自己的门派"&mp&"！</font>"		'聊天数据
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
%>
<html>
<head>
<title>〖<%=Application("aqjh_chatroomname")%>〗</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body background="../jhimg/bk_hc3w.gif" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font color="#FF0000" size="+3">弟子管理</font><br>
<br>
江湖告示：<%=message%> </div>
<hr>
<div align="center">版权所有『快乐江湖总站』</div> 
</body>
</html>