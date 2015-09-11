<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
select case Request.QueryString("rst")
case "0"
response.write "<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>按这里<a href='aqjh_reset.asp?rst=1'>给聊天室所有房间发送重启动告示</a><br><br> 请注意只有修改[聊天房间]参数后或者系统出现意外的错误时才需要[重启系统]"
response.end
case "1"
says="<font size=5 color=red><b>【系统重启】江湖将在3分钟后重新启动，请存点退出！</b></font>"
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
response.write "<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>通知发送完成!返回再次<a href='aqjh_reset.asp?rst=0'>通知</a>,按这里<a href='aqjh_reset.asp?rst=2'>重启动系统</a>"
response.end
case "2"
call Gohome()
Application("st_gohome")="go"
call Gouser()
Session("Gohome")="go"
case else
response.end
end select
sub Gohome
aqjh_data="../aqjh_data/aqjh.asp"
Application("aqjh_room")=""
Application("aqjh_npc")=""
Application("aqjh_usermdb")="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(aqjh_data)
'江湖秀数据库
Application("showmdb")="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../aqjh_data/show.asp")
Application("renwudb")="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../aqjh_data/renwu.asp")
Application("jhshowurl")="http://qqshow-item.tencent.com"    	'这是使用tencent的图片
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM r order by id",conn
 do while Not rs.Eof
Application("aqjh_room")=Application("aqjh_room")&rs("a")&"|"&rs("b")&"|"&rs("c")&"|"&rs("d")&"|"&rs("e")&"|"&rs("f")&"|"&rs("g")&"|"&rs("h")&"|"&rs("i")&"|"&rs("j")&"|"&rs("l")&"|"&rs("m")&";"
Application("aqjh_npc"&rs("a"))=""
	rs.MoveNext
 loop
Application("aqjh_copyright")="爱情江湖"
Application("aqjh_ver")="爱情江湖9.2"
Application("aqjh_softuser")=-63617
Application("aqjh_softjm")=-28972
Application("aqjh_tltie")="【"&Application("aqjh_chatroomname")&"】"&Application("aqjh_homepage")&"购买会员支持江湖发展("&Application("aqjh_user")&")"
rs.close
set rs=nothing 
conn.close
set conn=nothing
Dim nameindex(0)
aqjh_roominfo=split(Application("aqjh_room"),";")
for roomsn=0 to ubound(aqjh_roominfo)-1
	 Application("aqjh_onlinelist"&roomsn)=nameindex
	 fenroom=split(aqjh_roominfo(roomsn),"|")
	 application("aqjh_chatroomname"&roomsn)=fenroom(0)
	Application("aqjh_useronlinename"&roomsn)=""

next 

 Dim wbq(0)
 Application("aqjh_webicq")=wbq
 webicqname=" "
 Application("aqjh_webicqname")=webicqname
End sub
sub Gouser
 Session.Timeout=1
End Sub
sub Gouser
 Session.Timeout=1
End Sub
%>
<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>重启完毕<br><br><a href=../index.asp target=_top>重新登录</a>