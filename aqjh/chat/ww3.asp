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
if Application("aqjh_ww3")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ������Ʒ�ѱ���ȡ����ֻ������һ�ѳ���ʺ��');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww3"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('��ʾ���Ƿ��������ܾ�ִ�У�');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww3"))))

Application.Lock
Application("aqjh_ww3")=0
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
dim js(17)
js(0) ="��������"
js(1) ="��Ǯ��"
js(2) ="����޼"
js(3) ="��Ҷ��"
js(4) ="��Ӱ����"
js(5) ="���"
js(6) ="����"
js(7) ="÷����"
js(8) ="�佣"
js(9) ="�ɻ�ʯ"
js(10)="ͭǮ��"
js(11)="����"
js(12)="�ɵ�"
js(13)="����"
js(14)="͸�Ƕ�"
js(15)="������"
js(16)="������"
randomize()
myxy = Int(Rnd*17)
s=Int(Rnd*1000)
zstemp=add(rs("w4"),js(myxy),1)
conn.execute "update �û� set w4='"&zstemp&"' where ����='" & aqjh_name & "'"
kl="��<b>"&aqjh_name&"</b>ƽʱ���������ڵģ��кô���ʱ������죡����<b>"&js(myxy)&"</b>1�����ѱ����������</font>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<b>�����ᰵ����</b>"&kl

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