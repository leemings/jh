<%
username=session("Ba_jxqy_username")
biological=session("Ba_jxqy_userfight")
if username="" then 
	msg="top.location.replace('../error.asp?id=016');"
else
	if biological="none" then
		msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>������</FONT>û��ս������������Ч<br>');"
	else
		attname=Request.QueryString("attack")
		if instr(attname,"'")<>0 then 	Response.End 
		set conn=server.CreateObject("adodb.connection")
		conn.Open Application("Ba_jxqy_connstr")
		set rst=server.CreateObject("adodb.recordset")
		if Application("Ba_jxqy_fellow")=true then
			rst.Open "select tu.����,tu.����,ta.��������+tu.���� as ����,tu.����,ta.�������� from ���� ta inner join �û� tu on tu.����=ta.���� where ta.��ʽ='"&attname&"' and ta.����='"&username&"' and tu.��Ա=true",conn
		else
			rst.Open "select tu.����,tu.����,ta.��������+tu.���� as ����,tu.����,ta.�������� from ���� ta inner join �û� tu on tu.����=ta.���� where ta.��ʽ='"&attname&"' and ta.����='"&username&"'",conn
		end if		
		if rst.EOF or rst.BOF then
			msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>������</FONT>�Բ������޷�ʹ��"&attname&"���й���������������û��ѧ��"&attname&msg&"<br>');"
		elseif rst("��������")>rst("����") then
			msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>������</FONT>�Բ������޷�ʹ��"&attname&"���й���������������������<br>');"
		else
			fattack=rst("����")
			fdence=rst("����")
			fhp=rst("����")
			fmp=rst("����")
			ump=rst("��������")
			rst.Close
			rst.Open "select tf.hp,tb.attack,tb.defence,tb.encourage from fight tf inner join biological tb on tb.biological=tf.biological where tf.biological='"&biological&"' and tf.username='"&username&"' and fnum<5",conn
			if rst.EOF or rst.BOF then 	
				msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>��ս�����밮��Ұ�����һ��ֻ�ܴ������</FONT><br>');parent.behfrm.location.href='action.asp';"
			else	
				thp=rst("hp")
				tattack=rst("attack")
				tdefence=rst("defence")
				encourage=rst("encourage")
				attack=fattack-tdefence
				if attack>10000 then 
					attack=10000-clng(rnd()*100)
				elseif attack<=10 then
					attack=clng(rnd()*100)
				end if	
				defence=tattack-fdefence
				if defence>10000 then
					defence=10000-clng(rnd()*100)
				elseif defence<=10 then
					defence=clng(rnd()*100)
				end if
				fhp=fhp-defence
				fmp=fmp-ump
				thp=thp-attack
				if thp>0 and fhp>0 then
					conn.Execute "update �û� set ����="&fhp&",����="&fmp&" where ����='"&username&"'"
					conn.Execute "update fight set hp="&thp&" where username='"&username&"'"
					msg="parent.confrm.document.form1.move.disabled=true;parent.msgfrm.document.writeln('<FONT color=#ff0000>��ս����</FONT>"&username&"ʹ��"&attname&"����"&biological&"��ʹ֮����������"&attack&"��<br><FONT color=#ff0000>��ս����</FONT>"&username&"�⵽"&biological&"��������������"&defence&"��<br>');parent.confrm.document.form1.hp.value='������"&fhp&"';parent.confrm.document.form1.mp.value='������"&fmp&"';parent.behfrm.form1.hp.value='����:"&thp&"';"
				elseif fhp>0 and thp<0 then
					staple_M=mid(encourage,1,2)
					staple_C=mid(encourage,4)
					select case staple_M
						case "����"
							conn.Execute "update �û� set ����="&fhp&",����="&fmp&",����=����+"&clng(staple_c)&" where ����='"&username&"'"
						case "����"
							conn.Execute "update �û� set ����="&fhp&",����="&fmp&",����=����+"&clng(staple_c)&" where ����='"&username&"'"
						case "��Ʒ"
							rst.Close 
							rst.Open "select * from ��Ʒ where ������='"&username&"' and ����='"&staple_C&"'",conn
							if rst.EOF or rst.BOF then
								rst.close 
								rst.Open "select * from �̵� where ����='"&staple_C&"'",conn
								conn.Execute "insert into ��Ʒ(����,����,����,����,����,����,��Ч,����,�۸�,������,����,װ��) values('"&rst("����")&"','"&rst("����")&"',"&rst("����")&","&rst("����")&","&rst("����")&","&rst("����")&",'"&rst("��Ч")&"',1,"&rst("�۸�")\2&",'"&username&"',false,false)"
							else
								conn.Execute "update ��Ʒ set ����=����+1 where ������='"&username&"' and ����='"&staple_C&"'"
							end if
							conn.Execute "update �û� set ����="&fhp&",����="&fmp&" where ����='"&username&"'"
					end select
					conn.Execute "update fight set fnum=fnum+1 where username='"&username&"'"
					session("Ba_jxqy_userfight")="none"
					msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>��ս����</FONT>"&username&"ʹ��"&attname&"����"&biological&"��ʹ֮�������õ�"&staple_M&staple_C&"<br><FONT color=#ff0000>��ս����</FONT>"&username&"�⵽"&biological&"������������������"&defence&"��<br>');parent.confrm.document.form1.hp.value='������"&fhp&"';parent.confrm.document.form1.mp.value='������"&fmp&"';parent.confrm.document.form1.move.disabled=false;parent.behfrm.location.href='action.asp';"
				elseif fhp<=0 then
					logintime=dateadd("s",1800,now())
					logintimetype="#"&month(logintime)&"/"&day(logintime)&"/"&year(logintime)&" "&hour(logintime)&":"&minute(logintime)&":"&second(logintime)&"#"
					conn.Execute "update �û� set ����=0,״̬='����',����¼ʱ��="&logintimetype&" where ����='"&username&"'"
					session.Abandon
					msg="top.location.replace('../error.asp?id=054');"
				end if
			end if	
		end if
		rst.Close
		set rst=nothing
		conn.Close
		set conn=nothing	
	end if
end if
Response.Write "<script language=javascript>"&msg&"</script>"
%>