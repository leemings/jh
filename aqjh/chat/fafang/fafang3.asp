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
if aqjh_grade<9 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����һЩֻ��9������ſ��Բ���!');window.close();}</script>"
	Response.End 
end if
'#####���䴦��#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
if Instr(LCase(useronlinename),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(0)<>"�������" then
 Response.Write "<script language=javascript>{alert('��ʾ��������Ʒ��ȥ������');}</script>"
 Response.End
end if
cz=LCase(trim(Request.QueryString("cz")))
value=int(abs(clng(Request.QueryString("value"))))
randomize
s=int(rnd*value)+1
yn=0
select case cz
case "����"
	yn=0
        Application.Lock
	Application("aqjh_yb")=s
	Application.UnLock
	abc="<a href='fafang/jk.asp?tl="&Application("aqjh_yb")&"' target='d'><img src='img/jk.GIF' border='0'></a>"
	fafang="<bgsound src=wav/py.mid loop=1><font color=red>�����𿨡���ҿ�������ʲôѽ����,ԭ���ǽ𿨡�"&s&"�顱!��ҿ�����,˭������˭��!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "�����"
	yn=0
        Application.Lock
	Application("aqjh_yb")=s
	Application.UnLock
	abc="<a href='fafang/jb.asp?tl="&Application("aqjh_yb")&"' target='d'><img src='img/jinbi.GIF' border='0'></a>"
	fafang="<bgsound src=wav/py.mid loop=1><font color=red>������ҡ���ҿ�������ʲôѽ����,ԭ����<font color=blue>"&aqjh_name&"</font>���򽭺��Ľ�ҡ�"&s&"ö��!��ҿ�����,˭������˭��!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "�����"
	yn=0
	Application.Lock
	Application("aqjh_yb")=s
	Application.UnLock
	abc="<a href='fafang/cd.asp?tl="&Application("aqjh_yb")&"' target='d'><img src='img/cd.GIF' border='0'></a>"
	fafang="<bgsound src=wav/lw.mid loop=1><font color=red>������㡿</font><font color=red><font color=blue>"&aqjh_name&"</font>���ո��ˣ�����㡰"&s&"�㡱�������!˭������˭�ģ�</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "������"
	yn=0
	Application.Lock
	Application("aqjh_yb")=s
	Application.UnLock
	abc="<a href='fafang/dd.asp?tl="&Application("aqjh_yb")&"' target='d'><img src='img/dd.GIF' border='0'></a>"
	fafang="<bgsound src=wav/MYFAVOR.mid loop=1><font color=red>�������㡿</font><font color=red><font color=blue>"&aqjh_name&"</font>���ո��ˣ������㡰"&s&"�㡱�������!˭������˭�ģ�</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "�ݶ�����"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>�������ӡ�</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>�������Ӵ�Ժ�����������һ������Ķ��ӣ�Ҫȥ�ϼ�ȥ��˭֪����©��һ·�϶��ӵ���<font color=blue>+"& s &"</font>�����ٺ١��������зݣ�</font>"
end select
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
if yn=1 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("aqjh_usermdb")
	for nowinroom=0 to chatroomnum
		useronlinename=Application("aqjh_useronlinename"&nowinroom)
		online=Split(Trim(useronlinename)," ",-1)
		x=UBound(online)
		for i=0 to x
			conn.Execute "update �û� set "&cz&"="&cz&"+" & s & " where ����='" & online(i) & "'"
		next
	next
	conn.close
	set conn=nothing
end if
says=fafang
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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