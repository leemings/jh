<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'ͨ������wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_tjf then
	Response.Write "<script language=JavaScript>{alert('��ʾ��ͨ������Ҫ["&grade_tjf&"]�����Ϲ���Ա������');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
if mycz="/���" then
	says="<font color=green>�����ͨ����</font><font color=" & saycolor & ">"+tongji(mid(says,i+1),mycz)+"</font>"
else
	says="<font color=green>��ͨ���˷���</font><font color=" & saycolor & ">"+tongji(mid(says,i+1),mycz)+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'ͨ����
function tongji(fn1,mycz)
fn1=trim(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����,grade,ͨ�� from �û� where ����='" & fn1 &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������������û��������ڣ�');}</script>"
	Response.End
end if
if sjjh_grade<=rs("grade") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�����ԶԸ߼�����Ա������');}</script>"
	Response.End
end if
if mycz="/���" then
	if rs("ͨ��")<>True then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ:��Ҳû�б�ͨ��,�޷�������');}</script>"
		Response.End
	end if
	conn.execute "update �û� set ͨ��=False,����=True where ����='" & fn1 &"'"
	tongji= "����Ա:<font color=red>����"&fn1&"�Ѿ����ٸ��������״̬��������������Ľ��������ǵ�ͷ��~~"
else
	if rs("ͨ��")<>False then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ:���Ѿ���ͨ��,�޷�������');}</script>"
		Response.End
	end if
	conn.execute "update �û� set ͨ��=True,����=False where ����='" & fn1 &"'"
	tongji= "����Ա:<font color=red>����"&fn1&"�»����������в��壬�����ѱ��ٸ�ͨ����ɱ��ͨ�����Ĵ������õ�100��~~"
end if
conn.execute "insert into l(b,c,d,a,e) values ('" & sjjh_name & "','" & fn1 & "','����',now(),'������Ա"&mycz&"')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>