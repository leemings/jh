<%
username=session("Ba_jxqy_username")
biological=session("Ba_jxqy_userfight")
if biological="" or biological="none" then
	msg="<FONT color=#ff0000>������</FONT>û������ɹ��㲶׽<br>"
elseif username="" then	
	msg="<FONT color=#ff0000>������</FONT>��û�е�¼��ʱ�Ͽ�����,�����½���<br>"
else
	randomize()
	nowtime=now()
	nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("Ba_jxqy_connstr")
	set rst=server.CreateObject("adodb.recordset")
	if Application("Ba_jxqy_fellow")=true then
		rst.Open "select ����,���� from �û� where ����='"&username&"' and ����>=10 and ��Ա=true",conn
	else
		rst.Open "select ����,���� from �û� where ����='"&username&"' and ����>=10",conn
	end if
	if rst.EOF or rst.BOF then
		msg="<FONT color=#ff0000>������</FONT>��������,����Ϣһ�������������İ���<br>"
	else	
		fhp=rst("����")
		fdefence=rst("����")
		rst.Close
		rst.Open "select * from biological where biological='"&biological&"'",conn
		maxsinew=rst("hp")
		encourage=rst("encourage")
		encourage_m=Mid(encourage,1,2)
		encourage_c=Mid(encourage,4)
		attack=rst("attack")-fdefence
		if attack>10000 then
			attack=10000-clng(rnd()*100)
		elseif attack<100 then
			attack=clng(rnd()*1000)
		end if
		fhp=fhp-attack
		if fhp<=0 then
			conn.Execute "update �û� set ״̬='����',����¼ʱ��="&nowtimetype&" where ����='"&username&"'"
			session.Abandon
			Response.Write "<script language=javascript>top.location.replace('../error.asp?id=054');</script>"
			msg="<FONT color=#ff0000>������</FONT>�㱻"&biological&"������!<br>"
		else
			conn.Execute "update �û� set ����="&fhp&",����=����+1,����=����-10 where ����='"&username&"'"
			rndcatch=rnd()*99+1
			if rndcatch<15 then
				rst.Close
				rst.Open "select * from pet where username='"&username&"' and exist=true"
				if rst.EOF or rst.BOF then
					conn.Execute "insert into pet(username,biological,sinew,maxsinew,cleanliness,health,happy,compete,exist,encourage_m,encourage_c,option_m,option_t) values('"&username&"','"&biological&"',"&maxsinew*rndcatch\100&","&maxsinew&",10,10,10,false,True,'"&encourage_m&"','"&encourage_c&"','��',"&nowtimetype&")"
					msg="<FONT color=#ff0000>����׽��</FONT>��׽�ɹ�,���Ժ�úø���!"&biological&"�ķ���ʹ���������½�"&attack&"<br>"	
				else
					msg="<FONT color=#ff0000>����׽��</FONT>��׽ʧ��,���޷�ͬʱӵ����������!"&biological&"�ķ���ʹ���������½�"&attack&"<br>"	
				end if
			else
				msg="<FONT color=#ff0000>����׽��</FONT>��׽ʧ��,"&biological&"�ķ���ʹ���������½�"&attack&"<br>"	
			end if
		end if	
	end if
	rst.Close
	set rst=nothing
	conn.Close
	set conn=nothing
end if
Response.Write "<script language=javascript>parent.msgfrm.document.writeln('"&msg&"');</script>"
%>