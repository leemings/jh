<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'��������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����͵Ǯ��');}</script>"
	Response.End
end if
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
'�����뿪������Ѩ�ж�
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
says="<font color=green>��ħ��ʦ������������<font color=" & saycolor & ">"+qitaoshu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��������
function qitaoshu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����,��Ա�ȼ�,����,����,grade,�ݶ�����,ְҵ,���� from �û� where ����='" & aqjh_name &"'" ,conn,2,2
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫʹ����������!');}</script>"
	Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ�����ǳ����˰�������ɶ!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("ְҵ")<>"ħ��ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ȥְҵת��Ϊħ��ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
f=Minute(time())
if f<00 or f>25 then
	Response.Write "<script language=JavaScript>{alert('�����ж�������Ķ���ÿСʱ��ǰ25���ӣ�����������ʱ�䣡');window.close();}</script>"
	Response.End 
end if
fla=rs("����")
if fla<800000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��800000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select ����,�ȼ�,����,��Ա�ȼ�,����,grade,���� from �û� where ����='" & to1 &"'" ,conn,2,2
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��������������벻Ҫ����ʹ������!');}</script>"
	Response.End
end if
if rs("�ȼ�")<=30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���뽭�����ܶ�������');}</script>"
	Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('����Գ�������ʲô��������!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("grade")>=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ�Թ���Ա�������ҿ���Ҫ��̫���ˣ�����');}</script>"
	Response.End
end if
rs.close
rs.open "select �Ա�,���� FROM �û� WHERE ����='" & to1 &"'",conn
xb=rs("�Ա�")
yin=int(rs("����")/10)
if xb<>"Ů" then
	Response.Write "<script language=JavaScript>{alert('��������ֻ��Ů�������ã�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<10000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ûǮ�˲��ܶ�������');}</script>"
	Response.End
end if
conn.execute "update �û� set ����=����-"&yin&" where ����='" & to1 &"'"
conn.execute "update �û� set �ݶ�����=�ݶ�����+5,����=����+"&yin&" WHERE  ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����-500 WHERE  ����='" & aqjh_name &"'"
qitaoshu=aqjh_name & "<bgsound src=wav/Phant006.wav loop=1>��������ض�" & to1 & "˵����λƯ���ɰ���������õ�СMM<img src='img/79.gif' width='50' height='40'>�ɷ��ݽ�һ��Ǯ�����ã�������...˵��<font color=red>" & to1 & "MM</font>������˼�ش������ó�����֮һ�������ݸ���" & aqjh_name & "˵,�ÿ����ĺ��ӣ���ȥ���Եİɣ����û���. " & aqjh_name & "��˵õ�"&yin&"�����ӣ���������30000�㣬��������<font color=red>5</font>�㣬<font color=red>1000</font>�����������<font color=red>2</font>Ԫ��Ŷ~~�ٺ٣����˷���~~"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>�����ó�����֮һ�������ݸ���" & aqjh_name & "˵,�ÿ����ĺ��ӣ���ȥ���Եİɣ����û���. " & aqjh_name & "��˵õ�"&yin&"�����ӣ���������30000��"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>
