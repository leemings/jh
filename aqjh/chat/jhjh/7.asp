<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Buffer=true
Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
name=Session("aqjh_name")
nowinroom=session("nowinroom")
if name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&name&"'",conn,2,2
if rs("进化")<>"白钢人" then 
Response.Write "<script language=JavaScript>{alert('提示：你不是白钢人,不能进化成现代人了！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("金币")<15000 then 
Response.Write "<script language=JavaScript>{alert('提示：你的金币不够,进化到现代人需要15000金币！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if 

if rs("转生")<3 then 
Response.Write "<script language=JavaScript>{alert('提示：你的资格不够,进化到现代人需要转生3次！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
conn.execute"update 用户 set 进化='现代人',金币=金币-15000,内力加=内力加+70,体力加=体力加+70,武功加=武功加+70 where 姓名='"&name&"'"
attack="<font color=red><b>【快乐江湖时代进化】</b></font><font color=ff00ff>恭喜<b><font color=red>〖" & name & "〗</font></b>时代进化现代人成功,内力和体力、武功上限各加70,<font color=red>希望继续努力,早日进化到更高层,大家祝贺！</font></font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
says=attack 
says=replace(says,chr(39),"\'") 
says=replace(says,chr(34),"\"&chr(34)) 
act="消息" 
towhoway=0 
towho="大家" 
addwordcolor="660099" 
saycolor="008888" 
addsays="对" 
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
Response.Write "<script Language=Javascript>alert('恭喜您进化成现代人成功-获得各上限+70,希望继续努力！');location.href = 'index.asp';</script>"
Response.End 
rs.close
conn.close
set rs=nothing
set conn=nothing
%>

<%
Sub ErrALT(Str)
	Response.Write "<META http-equiv=Content-Type content=text/html;charset=gb2312><script>alert('" & Str &"');location.href='index.ASP';</script>"	
	Response.End
End Sub
%>