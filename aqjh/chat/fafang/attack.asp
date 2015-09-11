<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_grade<2 then
	Response.Write "<script language=JavaScript>{alert('提示：僵尸王释放只有管理2级以上者可释放');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 where 状态='僵尸王' and 身份='僵尸王' order by 姓名 ",conn
sf=rs("身份")
attackname=rs("姓名")
tl=rs("体力")
id=LCase(trim(Request.QueryString("id")))
value=int(abs(clng(Request.QueryString("value"))))
randomize timer
tl=int(rnd*value)+1
select case r
case 0
	yn=0
	Application.Lock
	Application("aqjh_kl")=tl+100000
	Application.UnLock
	abc="<a href='fafang/ks.asp?tsk="&Application("aqjh_kl")&"' target='d'><img src='img/gui.GIF' border='0'></a>" 
 	attack="<bgsound src=wav/gui.wav loop=3><font color=red>【释放僵尸王】" & aqjh_name & "看到僵尸被囚禁太久立刻注入大量体力给僵尸王然后把其中一只僵尸王释放出来，僵尸王" & attackname & "体力："&tl+100000&"<br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>" 
 end select 
if yn=1 then 
	useronlinename=Application("aqjh_useronlinename"&nowinroom) 
	online=Split(Trim(useronlinename)," ",-1) 
	x=UBound(online) 
	Set conn=Server.CreateObject("ADODB.CONNECTION") 
	conn.open Application("aqjh_usermdb") 
	for i=0 to x 

		conn.Execute "update 用户 set " & attackname & "=" & attackname & "+" & tl & " where 姓名='" & online(i) & "'"
	next 
	conn.close 
	set conn=nothing 
end if 
says=attack 
says=replace(says,chr(39),"\'") 
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
Response.Redirect "../../ok.asp?id=705"
%>