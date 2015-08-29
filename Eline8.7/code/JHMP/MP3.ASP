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
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
id=trim(request("id"))
you=trim(request("you"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 身份='掌门' and 姓名='"&sjjh_name&"'" ,conn
if rs.eof or rs.bof then	
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：你不是掌门！！');window.close();</script>"
		response.end
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
conn.execute "update p set c=c-1 where a='" & id & "'"
conn.execute "update 用户 set 门派='游侠',身份='弟子',grade=1 where 姓名='" & you & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>【逐出弟子】" & sjjh_name & "将自己弟子："& you &"赶出了自己门派！</font>"		'聊天数据
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
%>
<html>
<head>
<title>〖<%=Application("sjjh_chatroomname")%>〗</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body background="../jhimg/bk_hc3w.gif" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font color="#FF0000" size="+3">弟子管理</font><br>
<br>
江湖告示：<%=message%> </div>
<hr>
<div align="center">〖<%=Application("sjjh_chatroomname")%>〗 <a href="mailto:hc3w@163.com"><%=Application("sjjh_chatroomname")%></a> </div>
</body>
</html>