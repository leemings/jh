<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<5 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��ֻ�����Ųſ��Խ���������!');window.close();}</script>"
	Response.End 
end if
if session("advtime")<>"" then
if session("advtime")>now()-0.0006 then  Response.Redirect "../error.asp?id=490"
end if
session("advtime")=now()
cz=LCase(trim(Request.QueryString("cz")))
value=int(abs(clng(Request.QueryString("value"))))
randomize timer
s=int(rnd*value)+1
yn=0
select case cz
case "Ͱ��"
	yn=0
	Application.Lock
	Application("sjjh_wy")=s+5000
	Application.UnLock
	abc="<a href='wy.asp?tl="&Application("sjjh_wy")&"' target='d'><img src='img/wy.GIF' border='0' width='80' height='70'></a>"
	fafang="<font color=red>Ͱ��������ԭ�������������տ���Ŀ�ɿڴ���Ҳ��֪�����������ĳ����֡����ȴ�����˵�����Ų��ܴ�ร�Ͱ��������"&s+5000&"</font><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"
case "ľ��"
	yn=0
	Application.Lock
	Application("sjjh_cz")=s+5000
	Application.UnLock
	abc="<a href='cz.asp?tl="&Application("sjjh_cz")&"' target='d'><img src='img/cz.GIF' border='0' width='80' height='70'></a>"
	fafang="<font color=red>ľ���ִ�����ԭ�������������տ���Ŀ�ɿڴ���Ҳ��֪�����������ĳ����֡����ȴ�����˵�����Ų��ܴ�ร�ľ��������"&s+5000&"</font><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"
case "����"
	yn=0
	Application.Lock
	Application("sjjh_coco")=s+5000
	Application.UnLock
	abc="<a href='coco.asp?tl="&Application("sjjh_coco")&"' target='d'><img src='img/co.GIF' border='0' width='80' height='70'></a>"
	fafang="<font color=red>�����ִ�����ԭ�������������տ���Ŀ�ɿڴ���Ҳ��֪�����������ĳ����֡����ȴ�����˵�����Ų��ܴ�ร�����������"&s+5000&"</font><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"		
end select
if yn=1 then
	useronlinename=Application("sjjh_useronlinename"&nowinroom)
	online=Split(Trim(useronlinename)," ",-1)
	x=UBound(online)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("sjjh_usermdb")
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
Response.Redirect "../ok.asp?id=705"
%>

<head>
</head>
