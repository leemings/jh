<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
myid=Request.form("id")
name=request.form("name")
pass=trim(request.form("pass"))
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
if name="" or pass="" then
		Response.Write "<script Language=Javascript>alert('��ʾ���ǲ��ǲ��뱨��Ѫ���ˣ��������ͽ��ſ��������');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
pass=md5(pass)
rs.open "select * from �û� where id="& myid &" and ����='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����û�и����������˰���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if trim(pass)<>rs("����") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������벻�ԣ���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("����")>-100 and rs("״̬")<>"��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����û�и������˻�û��������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("״̬")<>"��" and rs("״̬")<>"����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����˱�ϵͳ����С���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	'��Ա�ȼ��ж�
	if rs("��Ա�ȼ�")>1 then
		conn.execute("update �û� set ״̬='����',����=1000,����=true where ����='"&name&"'")
		Response.Write "<script Language=Javascript>alert('OK,����2����Ա�������㸴����ʲôҲû�ж������ǣ����ǻ��ǲ�ϣ����������!');</script>"
	else
		conn.execute("update �û� set ״̬='����',����=1000,����=100,����=100,����=true where ����='"&name&"'")
			Response.Write "<script Language=Javascript>alert('OK,�������Ѿ������ˣ���Ҫ�����˰�,�����2�����ϻ�Ա�������κ���ʧ!');</script>"
	end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ϲ������ɹ��ĸ�����!��ס����һ��Ҫ����!');window.close();</script>"
says="<font color=red><b>��ԡ��������</b>["&name&"]</font><font color=blue>·������ʩ���������ٴλ����������ɱ�ҵ��������ˣ���һ��Ҫ�����Ǹ��𣡣���</font>"
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & "ϵͳ����" & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
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