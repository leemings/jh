<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=210"
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
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
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
conn.open Application("aqjh_usermdb")
rs.open "select ����,���,���� from �û� where ����='"&aqjh_name&"'",conn
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
if to1="��" or to1="��" or to1="δ��" or to1="δ��" then Response.Redirect "../error.asp?id=130"
if left(to1,3)="%20" OR InStr(to1,"=")<>0 or InStr(to1,"`")<>0 or InStr(to1,"'")<>0 or InStr(to1," ")<>0 or InStr(to1,"��")<>0 or InStr(to1,"'")<>0 or InStr(to1,chr(34))<>0 or InStr(to1,"\")<>0 or InStr(to1,",")<>0 or InStr(to1,"<")<>0 or InStr(to1,">")<>0 then Response.Redirect "../error.asp?id=120"
if to1=Application("aqjh_automanname") or to1="���" or Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܴ�λ�������˻��ң�������Ҫ��λ���˲����������');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if pass="" then	Response.Redirect "../../error.asp?id=432"
if aqjh_name=to1 then
	Response.Write "<script Language=Javascript>alert('���Ѿ��������ˣ������Լ���ʲôλѽ���ǲ��Ǻ����ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
%><!--#include file="../../pass.asp"--><%
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
	Response.Write "<script Language=Javascript>alert('���ǹٸ���Ա���Ѿ�������,���ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<>pai or rs("�ȼ�")<35 then
	Response.Write "<script Language=Javascript>alert('���ĵȼ�����35���������������µ��ӣ�������λ��');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
tmprs=conn.execute("Select count(*) As ���� from �û� where �ȼ�>=18 and ������='"& to1 &"'")
lr=tmprs("����")
set tmprs=nothing
if lr<3 then
	Response.Write "<script Language=Javascript>alert('��ʾ���������˼�¼����3�ˣ������������˵ĵȼ���û�е�18�����ϣ�');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "select ��Ա�ȼ� from �û� where ����='" & to1 &"'",conn,2,2
hy=rs("��Ա�ȼ�")
if hy=rs("��Ա�ȼ�")and (hy=1 or hy=2 or hy=3 or hy=4 or hy=5) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�Է��ǵȼ���Ա������λ����,��ʲô��������վ��˵ȥ!');window.close();}</script>"
		response.end
end if
mess="<font color=red>"&pai&"����"&aqjh_name&"�ֽ�����֮λ���������µ���"&to1&"!</font></marquee>"
conn.execute "update �û� set ���='����',grade=1,��Ǯ=now() where ����='" & aqjh_name & "'"
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
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ����������λ�ɹ������˳����½��뽭����');window.close();</script>"
response.end
%>