<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
if regjm<>regjm1 then
	%>
	<script language=vbscript>
	alert ("���������֤�벻��ȷ��Ӧ������:<%=regjm%>")
	location.href = "javascript:history.back()"
	</script>
	<%
	response.end
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	%>
	<script language="vbscript">
	 alert "�㲻�ܽ��в��������д˲���������������ң�"
	 window.close()
	</script>
	<%
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,���,���� from �û� where ����='"&sjjh_name&"'",conn
if rs("���")<>"����" then
	%>
	<script language="vbscript">
	 alert "���ֲ������ţ�����ʲô��"
	 window.close()
	</script>
	<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
pass=request.form("password")
pass=trim(pass)
to1=trim(request.form("xrzm"))
if instr(to1,"or")<>0 or instr(pass,"or")<>0 then Response.Redirect "../error.asp?id=54"
if to1="��" or to1="��" or to1="δ��" then Response.Redirect "../error.asp?id=130"
if left(to1,3)="%20" OR InStr(to1,"=")<>0 or InStr(to1,"`")<>0 or InStr(to1,"'")<>0 or InStr(to1," ")<>0 or InStr(to1,"��")<>0 or InStr(to1,"'")<>0 or InStr(to1,chr(34))<>0 or InStr(to1,"\")<>0 or InStr(to1,",")<>0 or InStr(to1,"<")<>0 or InStr(to1,">")<>0 then Response.Redirect "../error.asp?id=120"
if to1=Application("sjjh_automanname") or to1="���" or Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܴ�λ�������˻��ң�������Ҫ��λ���˲����������');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if pass="" then	Response.Redirect "../error.asp?id=432"
if sjjh_name=to1 then
	Response.Write "<script Language=Javascript>alert('���Ѿ��������ˣ������Լ���ʲôλѽ���ǲ��Ǻ����ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
%><!--#include file="../pass.asp"--><%
pass=md5(pass)
if rs("����")<>pass then
	%>
	<script language="vbscript">
	 alert "�Բ�������������"
	 window.close()
	</script>
	<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")="�ٸ�" then
	Response.Write "<script Language=Javascript>alert('��Ҫ��ʲô����');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
pai=rs("����")
rs.close
rs.open "SELECT ����,�ȼ�,���,grade FROM �û� WHERE trim(����)='" & to1 & "'",conn
if rs.bof or rs.eof then
	Response.Write "<script Language=Javascript>alert('��Ҫ��λ��˭ѽ�������к���û������ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
if rs("grade")>=6 or rs("����")="�ٸ�" then
	Response.Write "<script Language=Javascript>alert('���ǹٸ���Ա�������Զ�������������λ������');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<>pai or rs("�ȼ�")<50 then
	Response.Write "<script Language=Javascript>alert('���ĵȼ�����50���������������µ��ӣ�������λ��');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ��Ա�ȼ� FROM �û� WHERE ����='" & sjjh_name &"'",conn,3,3
hy=rs("��Ա�ȼ�")
if hy=rs("��Ա�ȼ�")and (hy<1) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��Ա�ȼ�����1��������λ����,��ʲô��������վ��˵ȥ!');window.close();}</script>"
		response.end
end if
mess="<font color=red>"&pai&"����"&sjjh_name&"�ֽ�����֮λ���������µ���"&to1&"!</font></marquee>"
conn.execute "update �û� set ���='����',grade=1,��Ǯ=now() where ����='" & sjjh_name & "'"
conn.execute "update �û� set ���='����',grade=5,��Ǯ=now() where ����='" & to1 & "'"
conn.execute "update p set b='"&to1&"' where a='"&pai&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>��С����Ϣ��</font><b><font color=green>"&mess&"</font></b>"			'��������
	says=replace(says,"'","\'")
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
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ����������λ�ɹ������˳����½��뽭����');window.close();</script>"
response.end
%>