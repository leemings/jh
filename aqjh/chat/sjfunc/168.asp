<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��������qqyh
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatroomsn=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>���������С�<font color=" & saycolor & ">"+qqyh()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function qqyh()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ����,ͨ��,����ʱ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
yinl=rs("����")
randomize timer
qjyl=int(rnd*yinl/100000)
if rs("ͨ��")=True then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	qqyh="[" & aqjh_name & "]" &":ͨ���������������У�Ҳ̫���˰ɣ�С��ץ��!"                                                 
	exit function
end if
if aqjh_jhdj<800 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	qqyh="[" & aqjh_name & "]" &":�㻹�ǽ���С������ƾ����ˮƽ�������У��������������ǵ�800��������!!"
	exit function
end if 
if  aqjh_grade>4 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	qqyh="[" & aqjh_name & "]" &":��������ݵ�λ���ˣ���ô��Ϊ����ô��С��ȥð��!!ֻ�г������µ��˲ſ���ȥ��!"
	exit function
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<=3000 then
	ss=3000-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	qqyh="["& aqjh_name & "]:�ٸ�����׽�����ٷ��أ���������ȥ�����������������ٹ�"& ss &"��������!"
	exit function
end if
randomize timer
r=Int(Rnd*7)
select case r
	case 0
		yinl=yinl+qjyl
		qqyh="["& aqjh_name & "]�����ק��һ��<img src='../chat/pic/273.gif'>,����������У����б���һ�����Լ������齣<img src='../hcjs/jhjs/images/14.jpg'>Ҳ̫����ˣ��ŵþ��ܣ������˱��������[" & aqjh_name &"]��������:"& qjyl
	case 1
		qqyh="["& aqjh_name & "]�����ק��һ�����齣<img src='../hcjs/jhjs/images/14.jpg'>����������У����б���һ���ϳ�<img src='../chat/pic/273.gif'>���ٺ٣���Ҳ̫����˰ɣ�[" & aqjh_name &"]ֻ�����־��ܣ�����������"&yinl&"ȫ�͸����ڣ���������֮�֡�"
		yinl=0
	case 2
		yinl=(yinl+int(qjyl/2))
		qqyh="["& aqjh_name & "]�����ק��һ�����齣<img src='../hcjs/jhjs/images/14.jpg'>����������У�����Ӧ���б��ڵĽ�Ӧ��[" & aqjh_name &"]��������:" & qjyl & "��һ��������б���"
	case 3
		yinl=yinl+5000000
		qqyh="["& aqjh_name & "]�������к����޴����֣�����ȴ�������ϵ��¸��ֽ�㣬����񵽱��˵��˵�����500��"
	case 4
		yinl=0
		qqyh="["& aqjh_name & "]�����ק��һ��<img src='../chat/pic/273.gif'>,����������У����б���һ�����Լ������齣<img src='../hcjs/jhjs/images/14.jpg'>Ҳ̫����ˣ��ŵþ��ܣ�һ��ȴ�ȵ������������"& "[" & aqjh_name &"]��������:" & qjyl & ",���ҵ��ǣ������Ѿ����ٸ���Ա��Χ�ˣ��Һû��ϵùٸ���Ա�������ϵ�Ǯ���Ϻ��������"
	case 5
		yinl=int(yinl/2)
		qqyh="["& aqjh_name & "]�������к����޴����֣�����ȴ�������ϵ��¸������ã��ó��Լ����ϵ�һ��Ǯ������֧�г���ȴ�����Ϳ���߳����У�֧���г���֪����ʲô��˼�����Ǳ�������¸�ľͰ�Ǯ�Ͻ�������"
	case else
		qqyh="["& aqjh_name & "]:���������Ѿ��°��ˣ�Ǯ�类�͵�������ȥ�ˣ�һ��ǮҲû�У���������ʲôѽ��"
end select
conn.execute "update �û� set ����="&yinl&",����ʱ��=now() where ����='" & aqjh_name &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>