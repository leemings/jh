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
if aqjh_grade<>5 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��ֻ�зǹٸ����Ųſ��Է�����!');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<3 then
	s=3-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����շŹ���,���ٵ�["&s&"]����������');location.href ='javascript:history.go(-1)';</script>"
	Response.End
end if
conn.Execute "update �û� set ����ʱ��=now() where ����='" & aqjh_name &"'"
set rs=nothing
conn.close
set conn=nothing
cz=LCase(trim(Request.QueryString("cz")))
value=int(abs(clng(Request.QueryString("value"))))
randomize timer
s=int(rnd*value)+1
yn=0
select case cz
case "Ͱ��"
	yn=0
	Application.Lock
	Application("aqjh_wy")=s+5000
	Application.UnLock
	abc="<a href='fafang/wy.asp?tl="&Application("aqjh_wy")&"' target='d'><img src='img/wy.GIF' border='0' width='80' height='70'></a>"
	fafang="<font color=green>����Ϣ��</font><font color=red>Ͱ�������˽������������տ���Ŀ�ɿڴ���Ҳ��֪�����������ĳ����֡����ȴ�����˵�����Ų��ܴ�ร�Ͱ��������"&s+5000&"(����:"&aqjh_name&")</font><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"
case "ľ��"
	yn=0
	Application.Lock
	Application("aqjh_cz")=s+5000
	Application.UnLock
	abc="<a href='fafang/cz.asp?tl="&Application("aqjh_cz")&"' target='d'><img src='img/cz.GIF' border='0' width='80' height='70'></a>"
	fafang="<font color=green>����Ϣ��</font><font color=red>ľ���ִ����˽������������տ���Ŀ�ɿڴ���Ҳ��֪�����������ĳ����֡����ȴ�����˵�����Ų��ܴ�ร�ľ��������"&s+5000&"(����:"&aqjh_name&")</font><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"
case "����"
	yn=0
	Application.Lock
	Application("aqjh_coco")=s+5000
	Application.UnLock
	abc="<a href='fafang/coco.asp?tl="&Application("aqjh_coco")&"' target='d'><img src='img/co.GIF' border='0' width='80' height='70'></a>"
	fafang="<font color=green>����Ϣ��</font><font color=red>�����ִ����˽������������տ���Ŀ�ɿڴ���Ҳ��֪�����������ĳ����֡����ȴ�����˵�����Ų��ܴ�ร�����������"&s+5000&"(����:"&aqjh_name&")</font><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"		
end select
if yn=1 then
	useronlinename=Application("aqjh_useronlinename"&nowinroom)
	online=Split(Trim(useronlinename)," ",-1)
	x=UBound(online)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("aqjh_usermdb")
	for i=0 to x
		conn.Execute "update �û� set "&cz&"="&cz&"+" & s & " where ����='" & online(i) & "'"
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