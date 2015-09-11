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
if aqjh_name="" then Response.Redirect "../error.asp?id=210"
if aqjh_grade<10 then
		Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');location.href = 'javascript:history.back()';}</script>"
		response.end
end if	
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	mess="官府的人喊到：吉时已到，罪大恶极的人犯"&Request("name")&"斩立决!!!!"
	if aqjh_grade=10 then
	conn.execute "delete * from 用户  WHERE 姓名='"&Request("name")&"'"
	set rs=nothing
	conn.close
	set conn=nothing
	end if
says="<font color=red>【消息】</font><font color=blank>"& mess &"</font>"			'聊天数据
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
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body bgcolor="#000000">
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr><td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="laofan.asp">返回</a></p>
</td></tr></table></td></tr></table></body>