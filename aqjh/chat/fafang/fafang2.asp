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
	Response.Write "<script language=JavaScript>{alert('��ʾ����ĵȼ�����,�㲻�ܲ���!');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 then
	s=5-sj
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
randomize
s=int(rnd*value)+1
yn=0
select case cz
case "����"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/gogo.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/dog.gif' border='0'></a>"
	fafang="<bgsound src=wav/dog.wav loop=2><font color=red>����Ϣ��ͻȻ��һ�����С���������ѽ�������������¹����˼�У���е���Ҫҧ�ң���Ҫҧ�ң������£�����һͷ�������������ң�����������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "˥��"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/boy.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/p42.gif' border='0'></a>"
	fafang="<bgsound src=wav/Ye150.wav loop=3><font color=red>����Ϣ��һ��˥����˵�����������Ů�ر�࣬�ܵ�����������Ů�����׼���ð��Ӵ�ѽ�������˹ٸ��н�����˥��������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "��ʿ"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/ws.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/bt.GIF' border='0'></a>"
	fafang="<bgsound src=wav/FX40A.wav loop=2><font color=red>����Ϣ���ۣ�����������һ����ʿ��Ҫ�Ҵ�ұ��䣬����˭�����ԣ���ʿ������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "�˹�����"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/lj.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/lj.GIF' border='0'></a>"
	fafang="<bgsound src=wav/zha.wav loop=2><font color=red>����Ϣ���˹����������й���������Ů��ɱ����ѽ����ѽ~~~~~~~~~~~~�˹�����������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "�˷�"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/rf.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/rf.GIF' border='0'></a>"
	fafang="<bgsound src=wav/FAQIAN.wav loop=1><font color=#FF6600>����������׽���˷���:Ҷ����(ye sanniang)�����ɶ���˷��ӣ������ǿ����ĺ����յ�����ȥ����Ҫ�����Ҳ����Ļ����Ҿͼ�һ������ɱһ�������˷���������+"&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "С͵"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/piaoke.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/piaoke.GIF' border='0'></a>"
	fafang="<bgsound src=wav/FAQIAN.wav loop=1><font color=#FF6600>����������ץС͵����ʼ������ͷ����Ҳ��ƽ��������Ȼ��λ��ͷ���Ե�С͵���뽭����˭������ס������л����</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "��Ů"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/mm.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/mm.GIF' border='0'></a>"
	fafang="<bgsound src=wav/tw.wav loop=2><font color=red>����Ϣ���׹�~~~~~~��������λƯ��������С�㡰��ҡ����ͺú�ҡ������ҡ�İ���ôҡ����ôҡ��������Ů������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "�̻�"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/yianhua.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/yianhua.GIF' border='0'></a>"
	fafang="<bgsound src=wav/yianhua.wma loop=1><font color=red>����Ϣ���׹�~~~~~~�����̻����ݣ��쿴ѽ���Ǻǣ����̻���Ҫ����¥����Ҫ������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/yianhua.wma loop=1>"&abc&"</marquee>"
case "ʨ��"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/xu.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/xu.GIF' border='0'></a>"
	fafang="<bgsound src=wav/shizi.wav loop=1><font color=red>����Ϣ���׹�~~~~~~����������ʨ�������ˣ���Ҳ��������ʨѽ��������Ҫ������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/xu.wma loop=1>"&abc&"</marquee>"
case "�ǵ�"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/fd.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/fd.GIF' border='0'></a>"
	fafang="<bgsound src=wav/feidian.wav loop=1><font color=red>����Ϣ�����۹�����ѷǵ�����˽�����Ϊ���Ϸ��ǵ䣬����Ŭ��ɱ������Ⱦ����"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "����"
yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='tuhu.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/jiao.gif' border='0'></a>"
    fafang="<bgsound src=wav/tufei.wav loop=1><font color=red>����Ϣ��ͻȻһȺ���˴��뽭�����٣�������ǿ�ȥ�˷˰�������������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "ƻ��"
	yn=0
	Application.Lock
	Application("aqjh_yb1")=s+100
	Application.UnLock
	abc="<a href='fafang/apple.asp?tl="&Application("aqjh_yb1")&"' target='d'><img src='img/apple.gif' border='0'></a>"
	fafang="<bgsound src=wav/baguo.wav loop=1><font color=red>����Ϣ��</font><font color=red>ͻȻ������ϵ�����һ����ƻ�����������ӣ�"&s+100&"����![����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "����"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/shu.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/shu.gif' border='0'></a>"
	fafang="<bgsound src=wav/eshu.wav loop=1><font color=red>����Ϣ��</font><font color=red>ͻȻ�佭������һֻ��������󡭡�������ѽ������������������˼�У���е���Ҫҧ�ң���Ҫҧ�ң������£�����һͷ���󴳽������ң�����������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "����"
yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='DABIAN.ASP?tl="&Application("aqjh_kl")&"' target='d'><img src='img/ying1.gif' border='0'></a>"
    fafang="<bgsound src=wav/niao.wav loop=2><font color=red>����Ϣ��һֻ�����ڤ���Ƿ��ٽ��������ң�̫���񰡣�һ�����кö����ģ���ҿ��!��������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100%  scrollamount=38>"&abc&"</marquee>"
case "����"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/wn.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/wn.GIF' border='0'></a>"
	fafang="<bgsound src=wav/tw.wav loop=2><font color=red>����Ϣ������ΪŮ�����Ǵ�������˧�������������������ǵ�˫�֡�ҡ����~~~~~~��������������"&s&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "��å"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s+100
	Application.UnLock
	abc="<a href='fafang/liumang.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/liumang.gif' border='0'></a>"
	fafang="<bgsound src=wav/liumang.wav loop=1><font color=red>����Ϣ��һ�������˵�����������Ů�ر�࣬�ܵ�����������Ů�����׼���ð��Ӵ�ѽ�������˹ٸ��н�������å���������"&s+100&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "�ϻ�"
	yn=0
	Application.Lock
	Application("aqjh_dalie")="�ϻ�"
	Application.UnLock
	fafang="<bgsound src=wav/hu.wav loop=1><font color=red>����Ϣ��</font><font color=red>ͻȻһֻҰ��<img src=img/laohu.gif>���뽭�������ˣ�������ǿ�ȥ���԰���[����:"&aqjh_name&"]</font>"
case "��ʬ"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s+100
	Application.UnLock
	abc="<a href='fafang/kl.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/gui.GIF' border='0'></a>"
	fafang="<bgsound src=wav/gui.wav loop=3><font color=red>����Ϣ��ͻȻ���������󡭡�����ʬѽ��һ��Ů�ӷ���һ���ֲ��ļ�У�һ����ʬ���������ң���ʬ������"&s+100&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "��ǹ"
	yn=0
	Application.Lock
	Application("aqjh_qi")=s+100
	Application.UnLock
	abc="<a href='fafang/qi.asp?tl="&Application("aqjh_qi")&"' target='d'><img src='img/tank.GIF' border='0'></a>"
	fafang="<bgsound src=wav/dz.wav loop=8><font color=red>����Ϣ��</font><font color=red>һ�Ѿ�����ǹ������������ԭ������ʿ����Ŀ�ɿڴ���Ҳ��֪������ʲô���������ȴ�����˵��������ǹ������"&s+100&"[����:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"
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
	for nowinroom=0 to chatroomnum
		saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
		addmsg saystr
	next
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