<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
id=request("id")
if not(isnumeric(id)) then
	Response.Write "<script language=JavaScript>{alert('����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,�Ա�,��ż,����,���,��Ա��,����,hw from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�ǽ������ˣ��������ã�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("�Ա�")="Ů" and (rs("����")<>"����" or rs("hw")<>"") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language=javascript>{alert('���ѻ��л�����С��������Ҫ�����ⷿ�����ǿ��ȥ�չ���ĺ��Ӱɡ���');location.href='javascript:history.go(-1)';}</script>"
	response.end
end if
yinl=rs("����")
jb=rs("���")
hf=rs("��Ա��")
myid=rs("id")
poxm=rs("��ż")
rs.close
if poxm="" or poxm="��" then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language=javascript>{alert('�㻹δ��飬���겻��δ�����Ͽ��ţ�');location.href='javascript:history.go(-1)';}</script>"
	response.end
end if
rs.open "select id,�Ա�,��ż,����,hw from �û� where ����='"&poxm&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�����ż���ǽ������ˣ��������ã�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("�Ա�")="Ů" and (rs("����")<>"����" or rs("hw")<>"") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language=javascript>{alert('��������ѻ��л�����С��������Ҫ�����ⷿ�����ǿ��ȥ�չ�����޶���С�ɡ���');location.href='javascript:history.go(-1)';}</script>"
	response.end
end if
poid=rs("id")
rs.close
'rs.open "select * from fw where yhid="&myid&" or yhid="&poid,conn
'if not(rs.eof or rs.bof) then
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script language=JavaScript>{alert('�����������Ѿ����Ϸ��䣬�������ޣ����������ã�');location.href = 'javascript:history.go(-1)';}</script>"
'	Response.End
'end if
'rs.close
rs.open "select * from fw where id="&id,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����С���в�û��������ݣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
yhid=rs("yhid")
if yhid<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�˷��������������ˣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
gmfs=rs("jgfs")
jg=rs("b")
fwm=rs("a")
rs.close
select case gmfs
	case "����"
		if jg>yinl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('���ô˷���һ����Ҫ��"&gmfs&"��"&jg&"�����Ǯ������');location.href = 'javascript:history.go(-1)';}</script>"
			Response.End
		end if
	case "���"
		if jg>jb then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('���ô˷���һ����Ҫ��"&gmfs&"��"&jg&"�����Ǯ������');location.href = 'javascript:history.go(-1)';}</script>"
			Response.End
		end if
	case "��Ա��"
		if jg>hf then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('���ô˷���һ����Ҫ��"&gmfs&"��"&jg&"�����Ǯ������');location.href = 'javascript:history.go(-1)';}</script>"
			Response.End
		end if
end select
conn.execute "update fw set yhid="&myid&",zysj=now() where id="&id
conn.execute "update �û� set "&gmfs&"="&gmfs&"-"&jg&" where ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
if Instr(Application("aqjh_useronlinename"&session("nowinroom"))," "&poxm&" ")<>0 then
says="<font color=red>������С����</font><font color=#8000FF>"&aqjh_name&"�ڰ��齭��������С��Ϊ����"&poxm&"������<font color=red>��"&fwm&"��</font>���䣬"&aqjh_name&"����С���еȺ�"&poxm&"��ȥѽ���Ǻ�..</font>"		'��������
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
end if
%>
<html>
<head>
<title>���÷��ݳɹ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#8800FF" text="#FFFF00">
<div align="center">
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>��ϲ��ɹ������˽�������С���е�[<b><font color="#00FFFF"><%=fwm%></font></b>]���䣬��ȥ֪ͨ��İ�����������۰ɡ� 
  </p>
  <p>&nbsp;</p>
  <p><a href="index.asp">����</a> 
  </p>
</div>
</body>
</html>