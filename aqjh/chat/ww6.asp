<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Application("aqjh_ww6")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ������Ʒ�ѱ���ȡ����ֻ������һ�ѳ���ʺ��');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww6"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('��ʾ���Ƿ��������ܾ�ִ�У�');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww6"))))

Application.Lock
Application("aqjh_ww6")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update �û� set ����=����-"&tempjs&" where ����='" & aqjh_name &"'"
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn
if rs("����")<-1000 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='���|����Ʒ������' where ����='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'���','����Ʒ������','����')"
	kl="<img src='pic/kl.gif'>̫�����ˣ�["&aqjh_name&"]�ڸ��������Ʒʱ����Ҵ���������"
	call boot(aqjh_name,"��ң������ߣ���ң�����������������������")
else
dim js(21)
js(0) ="��������"
js(1) ="������Զ"
js(2) ="����ŨŨ"
js(3) ="�Ҹ�ȪԴ"
js(4) ="�����Ҹ�"
js(5) ="Ũ���ƻ�"
js(6) ="�߷�����"
js(7) ="����ú�"
js(8) ="���鲻��"
js(9) ="��������"
js(10)="��ܰ�İ�"
js(11)="��������"
js(12)="����ɴ�"
js(13)="����Ǹ��"
js(14)="�Ҹ�����"
js(15)="��ɫ�赸"
js(16)="ˮ��֮��"
js(17)="������˵"
js(18)="����õ��"
js(19)="����ʱ��"
js(20)="�鶨һ��"
js(21)="��������"
randomize()
myxy = Int(Rnd*21)
s=Int(Rnd*1000)
s1=Int(Rnd*3)+1
zstemp=add(rs("w7"),js(myxy),s1)
conn.execute "update �û� set w7='"&zstemp&"' where ����='" & aqjh_name & "'"
kl="�ۣ���ô��ô���˰���<b>"&aqjh_name&"</b>�⻨���ܼ񵽣������⻨����MM˧�磬����û�в��Ϲ��ģ�������<b>"&js(myxy)&"</b>"&s1&"�䣬�������ˣ�"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<b>�������ʻ���</b>"&kl

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
%>
