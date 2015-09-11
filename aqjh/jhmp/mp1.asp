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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then	
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：江湖上找不到你的资料！');window.close();</script>"
		response.end
end if
if rs("门派")<>"" and  rs("门派")<>"无" and  rs("门派")<>"游侠"  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：要想加入门派请离开你现在的门派！！');window.close();</script>"
		response.end
end if
rs.close
rs.open "SELECT * FROM p WHERE a='" & id & "'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：江湖中没有这个门派！！');window.close();</script>"
		response.end
	end if
	tj=rs("g")
	sm=rs("e")
	rs.close
	rs.open "SELECT * FROM 用户 WHERE 姓名='"&aqjh_name&"'" &" and "&tj,conn
		if rs.eof or rs.bof  then
			rs.close
			set rs=nothing	
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('提示：加入["&id&"]的条件是："&sm&"！');window.close();</script>"
			response.end
		end if	
	message="你成功地加入了[" & id & "]这个门派!"
	conn.execute "update 用户 set 门派='" & id & "', 身份='弟子',grade=1 where 姓名='"&aqjh_name&"'"
	conn.execute "update p set c=c+1 where a='" & id & "'"
says="<font color=#ff0000>【加入门派】" & aqjh_name & "成功的加入了["& id &"]这个门派！</font>"			'聊天数据
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
<p align="center"><font color="#FF0000" size="+3">加入江湖门派</font></p>
<p align="center"><font color="#FFFFFF"><b><font color="#000000"><%=Application("aqjh_chatroomname")%>告示：<%=message%>
</font></b></font><font color="#000000"></font></p>
<hr>
</body>
</html>
