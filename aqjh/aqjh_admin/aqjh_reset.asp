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
response.write "<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>������<a href='aqjh_reset.asp?rst=1'>�����������з��䷢����������ʾ</a><br><br> ��ע��ֻ���޸�[���췿��]���������ϵͳ��������Ĵ���ʱ����Ҫ[����ϵͳ]"
response.end
case "1"
says="<font size=5 color=red><b>��ϵͳ��������������3���Ӻ����������������˳���</b></font>"
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
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
response.write "<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>֪ͨ�������!�����ٴ�<a href='aqjh_reset.asp?rst=0'>֪ͨ</a>,������<a href='aqjh_reset.asp?rst=2'>������ϵͳ</a>"
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
'���������ݿ�
Application("showmdb")="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../aqjh_data/show.asp")
Application("renwudb")="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../aqjh_data/renwu.asp")
Application("jhshowurl")="http://qqshow-item.tencent.com"    	'����ʹ��tencent��ͼƬ
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM r order by id",conn
 do while Not rs.Eof
Application("aqjh_room")=Application("aqjh_room")&rs("a")&"|"&rs("b")&"|"&rs("c")&"|"&rs("d")&"|"&rs("e")&"|"&rs("f")&"|"&rs("g")&"|"&rs("h")&"|"&rs("i")&"|"&rs("j")&"|"&rs("l")&"|"&rs("m")&";"
Application("aqjh_npc"&rs("a"))=""
	rs.MoveNext
 loop
Application("aqjh_copyright")="���ֽ���"
Application("aqjh_ver")="���ֽ���9.2"
Application("aqjh_softuser")=-63617
Application("aqjh_softjm")=-28972
Application("aqjh_tltie")="��"&Application("aqjh_chatroomname")&"��"&Application("aqjh_homepage")&"�����Ա֧�ֽ�����չ("&Application("aqjh_user")&")"
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
<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>�������<br><br><a href=../index.asp target=_top>���µ�¼</a>