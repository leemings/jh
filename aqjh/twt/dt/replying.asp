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
username=aqjh_name
if username=Application("aqjh_askuser") then response.redirect "../../error.asp?id=476"
if username="" then response.redirect "../../error.asp?id=127"
reply=trim(Request.form("reply"))
reply=server.HTMLEncode(reply)
if trim(Request.form("reply"))="" then Response.Redirect "../../error.asp?id=475"
if reply=Application("aqjh_reply") then
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="SELECT 银两 FROM 用户 WHERE 姓名='"&aqjh_name&"'"
rs.open sql,conn,1,3
rs("银两")=rs("银两")+Application("aqjh_asksilver")
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
says="【答 题】<font color=balck>" & Application("aqjh_ask") & "？</font>正确答案是：<font color=red>[ " & Application("aqjh_reply") & " ] </font>〖"& username &"〗因此获得了<font color=green>"& Application("aqjh_asksilver") &"</font>两银子的奖励!"			'聊天数据
Application.Lock
Application("aqjh_talkarr")=newtalkarr
Application("aqjh_ask")=""
Application("aqjh_reply")=""
Application("aqjh_askuser")=""
Application.UnLock
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
says=replace(says,chr(13),"")
says=replace(says,chr(10),"")
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
else
response.redirect "../../error.asp?id=479"
response.end
end if
%>
<script LANGUAGE="JavaScript">
if(window!=window.top){top.location.href=location.href;}
window.close();
</script>