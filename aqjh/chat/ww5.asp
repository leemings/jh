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
if Application("aqjh_ww5")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ������Ʒ�ѱ���ȡ����ֻ������һ�ѳ���ʺ��');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww5"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('��ʾ���Ƿ��������ܾ�ִ�У�');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww5"))))

Application.Lock
Application("aqjh_ww5")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update �û� set ����=����-"&tempjs&" where ����='" & aqjh_name &"'"
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn
if rs("����")<-100 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='���|����Ʒ������' where ����='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'���','����Ʒ������','����')"
	kl="<img src='pic/kl.gif'>̫�����ˣ�["&aqjh_name&"]�ڸ��������Ʒʱ����Ҵ���������"
	call boot(aqjh_name,"��ң������ߣ���ң�����������������������")
else
dim js(18)
js(0) ="õ������"
js(1) ="һ������"
js(2) ="����˫��"
js(3) ="���鱼��"
js(4) ="�Ұ�����"
js(5) ="ƬƬ����"
js(6) ="����ͬ��"
js(7) ="��������"
js(8) ="�ҵ�ǣ��"
js(9) ="��������"
js(10)="Ũ���ƻ�"
js(11)="ǧ����˼"
js(12)="ҡͷ��õ��"
js(13)="������"
js(14)="���˶�"
js(15)="��������"
js(16)="������˵"
js(17)="�����õ��"
randomize()
myxy = Int(Rnd*18)
s=Int(Rnd*1000)
zstemp=add(rs("w7"),js(myxy),10)
conn.execute "update �û� set w7='"&zstemp&"' where ����='" & aqjh_name & "'"
kl="��<b>"&aqjh_name&"</b>ƽʱ���������ڵģ��кô���ʱ������죡�ʻ�<b>"&js(myxy)&"</b>10�䣬�ѱ����������"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<b>����ȡ�ʻ���</b>"&kl

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
