<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
id=trim(request("id"))
if instr(id,"��")<>0 then
		Response.Write "<script language=JavaScript>{alert('���ؾ��棬��Ҫ����');window.close();}</script>"
		Response.End
end if
my=trim(request("my"))
if my<>aqjh_name then
		Response.Write "<script language=JavaScript>{alert('���ؾ��棬��Ҫ����');window.close();}</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� WHERE ���<>'����' and ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��ʾ���������Ҳ���������ϻ����������ţ�');window.close();</script>"
		response.end
end if
	if rs("����")="" or rs("����")="����" or rs("����")="��" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��ʾ���㲢�����ɣ���');window.close();</script>"
		response.end
	end if
	if rs("����")="����" or rs("����")="����"  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��ʾ�����ǳ����ˣ���������ɱ�ֲ������������뿪!����');window.close();</script>"
		response.end
	end if

if rs("��Ա�ȼ�")=0 then
	message="���뿪��" & rs("����") & "����ʧ�������书��������1000,��������������1000����"
	if rs("���")>1000 then
		conn.execute "update �û� set ���ɻ���=0,����=����+1,����='����',���='����', ����=����-1000,����=����-1000,�书=�书-1000,grade=1,����=1000,���=1000 where ����='"&aqjh_name&"'"
	else
		conn.execute "update �û� set ���ɻ���=0,����=����+1,����='����',���='����', ����=����-1000,����=����-1000,�书=�书-1000,grade=1,����=0 where ����='"&aqjh_name&"'"
	end if
else
	message="���뿪��" & rs("����") & "����Ϊ���ǽ�����Ա������ԭ����������Ǯ���䣡"	
	conn.execute "update �û� set ���ɻ���=0,����=����+1,����='����',���='����',grade=1 where ����='"&aqjh_name&"'"
end if
	conn.execute "update p set c=c-1 where a='" & id & "'"
says="<font color=#ff0000>��������ɡ�" & aqjh_name & "�ɹ����뿪��["& id &"]������ɣ����������ٵ�1000�����ɴ�����1�����ɹ��׼�Ϊ0��</font>"			'��������
says=replace(says,"'","\'")
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
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>��<%=Application("aqjh_chatroomname")%>��</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body background="../jhimg/bk_hc3w.gif" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font size="+3" color="#FF0000">�뿪����</font><br>
<br>
<br>
<%=Application("aqjh_chatroomname")%>��ʾ��<%=message%> </div>
<hr>
<div align="center">��Ȩ���С����ֽ�����վ��</div>
</body>
</html>
