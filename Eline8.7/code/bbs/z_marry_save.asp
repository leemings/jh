<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<%
dim content

toname=request.form("sBride")
content=request.form("sLoveWord")
guige=request.form("iStandard")
theyear=request.form("wed_b_y")
them=request.form("wed_b_m")
thed=request.form("wed_b_d")
theh=request.form("wed_b_h")
lasttime=request.form("sLastTime")
fumoney=request.form("cPayMethod")
total=guige*100+lasttime*100+400

call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�в鿴���������õ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	call marry_nav()
	call main()
end if
call activeonline()
call footer()

sub main()
	dim tousermoney,yourmoney
	set rs=server.createobject("adodb.recordset")
	'rs.open"select * from [user] where username='"&membername&"'",conn,3,1
	'yourmoney=rs("userwealth")
	'rs.close
	'dim tousermoney
	'yourmoney=conn.execute("select userwealth from [user] where username='"&membername&"'")(0)
	'tousermoney=conn.execute("select userwealth from [user] where username='"&toname&"'")(0)
	'if (fumoney=1 and total>yourmoney) or (fumoney=2 and ) then 
	'	Errmsg=Errmsg+"<br>"+"<li>�����ֽ��㣡"
	'	call dvbbs_error()
	'else
		rs.open"select * from [qiuhun] where (UserName='"&membername&"' or tousername='"&membername&"') and dlg=true",connj,3,1
		if rs.eof then
			rs.close
			Errmsg=Errmsg+"<br>"+"<li>�㻹û�����κ���������û�˴�Ӧ�����飡"
			call dvbbs_error()
			response.end
		elseif rs("jiehun") =true then
			rs.close
			Errmsg=Errmsg+"<br>"+"<li>���Ƕ��Ѿ��Ǽǽ���ˣ���ô�����Ǽǰ���"
			call dvbbs_error()
			response.end
		else
			if rs("username")=membername then
				ok=true
				name2=rs("tousername")
			elseif rs("tousername")=membername then
				ok=false
				name2=rs("username")
			end if
			rs.close
			if toname<>name2 then
				Errmsg=Errmsg+"<br>"+"<li>���ǲ����Ҵ����ˣ�"
				call dvbbs_error()
				response.end
			else
				yourmoney=conn.execute("select userwealth from [user] where username='"&membername&"'")(0)
				tousermoney=conn.execute("select userwealth from [user] where username='"&toname&"'")(0)
				if fumoney=1 then
					if yourmoney<total then
						ErrMsg=ErrMsg+"<br>"+"<li>�����ֽ��㣬��ѡ���������ʽ��"
						call dvbbs_error()
						response.end
					else
						conn.execute"update [user] set userwealth=userwealth-"&total&",wife='"&name2&"' where username='"&membername&"' " 
						conn.execute"update [user] set wife='"&membername&"' where username='"&name2&"' " 
					end if
				elseif fumoney=2 then
					if tousermoney<total then 
						ErrMsg=ErrMsg+"<br>"+"<li>�����˵��ֽ��㣬��ѡ���������ʽ��"
						call dvbbs_error()
						response.end
					else
						conn.execute"update [user] set userwealth=userwealth-"&total&",wife='"&membername&"' where username='"&name2&"' " 
						conn.execute"update [user] set wife='"&name2&"' where username='"&membername&"' " 
					end if
				elseif fumoney=3 then
					total=total/2
					if tousermoney<total or yourmoney<total then
						ErrMsg=ErrMsg+"<br>"+"<li>������һ���ֽ����ܽ���һ�룬��ѡ���������ʽ��"
						call dvbbs_error()
						response.end
					else
						conn.execute"update [user] set userwealth=userwealth-"&total&",wife='"&name2&"' where username='"&membername&"' " 
						conn.execute"update [user] set userwealth=userwealth-"&total&",wife='"&membername&"' where username='"&name2&"' " 
					end if
				end if
				set rs1=server.createobject("adodb.recordset")
				if ok then
					rs1.open"select * from [jie] where username='"&membername&"' and thename='"&name2&"'",connj,3,3
				else
					rs1.open"select * from [jie] where username='"&name2&"' and thename='"&membername&"'",connj,3,3
				end if
				rs1("long")=lasttime
				rs1("content")=content
				'rs1("year")=theyear&"-"&them&"-"&thed
				rs1("year")=now()
				rs1("type")=guige
				rs1("jiehun")=true
				rs1.update
				rs1.close
				connj.execute"update [qiuhun] set jiehun=true where username='"&membername&"' or tousername='"&membername&"'" 
				%>
				<script language="vbscript">
					msgbox"ף����Ǻǣ�����ˣ�����ȥ����������ѣ��ǵ�Ҫ׼ʱŶ",0,"<%=Forum_info(0)%>"
					location.href = "z_marry_index.asp"
				</script>
				<%
				set rs1 = Server.CreateObject("ADODB.recordset")
				rs1.open"Select * From [message]",conn,3,3
				rs1.addnew
				rs1("sender")="��������"
				rs1("incept")=name2
				rs1("title")="��ϲ"&membername&"�����ɷ���!"
				rs1("content")="��ϲ������"&now()&"��ɷ���<br>�뵽��������ȥ������"
				rs1("sendtime")=now()
				rs1("issend")=1
				rs1.update
				rs1.addnew
				rs1("sender")="����������"
				rs1("incept")=membername
				rs1("title")="��ϲ"&name2&"�����ɷ���!"
				rs1("content")="��ϲ������"&now()&"��ɷ���<br>�뵽��������ȥ������"
				rs1("sendtime")=now()
				rs1("issend")=1
				rs1.update
				rs1.close
			end if
		end if
	'end if
end sub %>