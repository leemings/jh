<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<%dim menu,body
menu=request.querystring("menu")%>
<html>
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	if menu="" then
 		call main()
	elseif menu=1 then
 	call manage()
	elseif menu=2 then
 	call addman()
 elseif menu=3 then
 	call savemod()
 elseif menu=4 then
 	call cleanvip()
 end if 	
end if%>

<%sub main() 
dim sql1,rs1,sql2,rs2 
sql1="select * from [vip]" 
set rs1=conn.execute(sql1) 
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
<form action=?menu=3 method=post>
<tr><th valign=middle height=23 align=left width="100%">VIP��Ա����ϵͳģʽ�趨</th></tr>
<tr><td valign=middle class=forumrow align=left width="75%">
˵����VIP��Ա����ϵͳ������ģʽ<br>
<font color="#FF0000">�Զ����ϵͳ</font>���û���ǰ̨���Կ�������������������������������룬�����ֱ�Ӽ�ΪVIP��Ա��VIP���ں���������ٴ���������ֱ�Ӽ�ΪVIP��Ա��<br>
<font color="#FF0000">�˹����ϵͳ</font>���û���ǰ̨���Կ���������������������������Եݽ����룬������ύ������Ա�����������׼���ΪVIP��Ա��VIP���ں���������������Եݽ��ٴ����룬������ύ������Ա�����������׼���ٴμ�ΪVIP��Ա��<br> 
<font color="#FF0000">���Զ����ϵͳ</font>���û���ǰ̨���Կ�������������������������������룬�����ֱ�Ӽ�ΪVIP��Ա��VIP���ں���������ٴ���������ֱ�Ӽ�ΪVIP��Ա��</td>
</tr><tr>
<td valign=middle class=forumrow align=left width="75%">
<p align="center">�Զ����VIP��Ա<input type=radio name="vipmod" value="0" <%if rs1("vipmod")=0 then%>checked<%end if%>>���������˹����VIP��Ա<input type=radio name="vipmod" value="1" <%if rs1("vipmod")=1 then%>checked<%end if%>>�����������Զ����VIP��Ա<input type=radio name="vipmod" value="2" <%if rs1("vipmod")=2 then%>checked<%end if%>>������<input type=submit name="submit" value="�� ��"></td></tr>
</form>
</table>
<br>

<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
<tr><th align=center colspan=4 width="95%" height=1><p align="left"><b>��̳VIP������������</b></p></th></tr> 
<form name="form" method="post" action="?menu=1"> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>�״�����VIP�����</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userwealth size="20" value="<%=rs1("wealth")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>�ٴ�����VIP�����</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userwealth2 size="20" value="<%=rs1("wealth2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>�״�����VIP���辭�飺</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userep size="20" value="<%=rs1("userep")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>�ٴ�����VIP���辭�飺</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userep2 size="20" value="<%=rs1("userep2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>�״�����VIP����������</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=usercp size="20" value="<%=rs1("usercp")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>�ٴ�����VIP����������</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=usercp2 size="20" value="<%=rs1("usercp2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>�״�����VIP����������</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userpower size="20" value="<%=rs1("power")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>�ٴ�����VIP����������</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userpower2 size="20" value="<%=rs1("power2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>�״�����VIP������������</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userart size="20" value="<%=rs1("article")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>�ٴ�����VIP������������</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=userart2 size="20" value="<%=rs1("article2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="170" height="25" class=forumrow>VIP��Ա��Ч�ڣ�(��)</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=vipdate size="20" value="<%=rs1("vipdate")%>"></td> 
	<td align=right width="170" height="25" class=forumrow>������ǰ֪ͨ��������(��)</td> 
	<td align=left width="180" height="25" class=forumrow><INPUT name=notifydays size="20" value="<%=rs1("notifydays")%>"></td> 
</tr> 
<tr> 
	<td align=center colspan=4 width="95%" height="27" class=forumrow><INPUT type=submit value=�޸� name=submit></td> 
</tr> 
</form> 
</table> 
<br> 
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder"> 
<tr><th align=center width="95%" height=1 colspan="2"><p align="left">��̳VIP��Ա����</p></th></tr>
<tr width="95%" > 
<td width="35%" class=forumrow><b>����VIP��Ա</b>������������ȷ��ΪVIP���û�������ȷ�ϴ��û��Ǳ���̳���û���</td>
<form name="form1" method="post" action="?menu=2"> 
<td width="65%" class=forumrow height="45">����Ƕ���û����ö��ŷָ���<br><INPUT name=vipname size="20">&nbsp; <INPUT type=submit value=ȷ�� name=submit></td> 
</form> 
</tr> 
<tr width="95%" > 
<td width="35%" class=forumrow><b>VIP��Ա����</b></td>
<form name="form1" method="get" action="z_admin_vipuser.asp">
<td width="65%" class=forumrow height="40"><INPUT type=submit value='�� �� �� ��' name=submit></td> 
</form>
</tr> 
<tr width="95%" > 
<td width="35%" class=forumrow><b>����VIP��Ա����</b></td> 
<form name="form1" method="get" action="z_admin_vipapplyuser.asp">
<td width="65%" class=forumrow height="40"><INPUT type=submit value='�� �� �� ��' name=submit></td> 
</form> 
</tr> 
<tr width="95%" > 
<td width="35%" class=forumrow><b>�����ѹ��ڵ�VIP��Ա</b></td> 
<form name="form1" method="get" action="?menu=4">
<td width="65%" class=forumrow height="40"><INPUT type=submit value='�� �� �� ��' name=submit></td> 
</form> 
</tr> 
</table> 
<br> 
<%end sub

sub savemod()
	dim rs,sql
	if request("vipmod")="" then
		response.write "ѡ�����Ϊ�գ�"
		response.end
	else
		set rs= server.createobject ("adodb.recordset") 
		sql = "select * from vip" 
		rs.Open sql,conn,1,3 
		if not(rs.eof and rs.bof) then 
			rs("vipmod")=request("vipmod") 
			rs.Update 
		end if
		rs.close
		set rs=nothing 
		response.write"�޸ĳɹ���" 
	end if
end sub

sub manage() 
	dim money,usercp,userep,userpower,userart,rs,sql,vipdate,money2,usercp2,userep2,userpower2,userart2,notifydays
	if not isInteger(request.form("userwealth")) or not isInteger(request.form("userart")) or not isInteger(request.form("userep")) or not isInteger(request.form("usercp")) or not isInteger(request.form("userpower")) or not isInteger(request.form("vipdate")) or not isInteger(request.form("userwealth2")) or not isInteger(request.form("userart2")) or not isInteger(request.form("userep2")) or not isInteger(request.form("usercp2")) or not isInteger(request.form("userpower2")) or not isInteger(request.form("notifydays")) then
		response.write "��������ݱ���Ϊ������"
		response.end 
	else 
		if request.form("userwealth")="" then 
			money=20
		else 
			money=int(request.form("userwealth")) 
		end if 
		if request.form("userep")="" then 
			userep=20 
		else 
			userep=int(request.form("userep")) 
		end if 
		if request.form("usercp")="" then 
			usercp=20 
		else 
			usercp=int(request.form("usercp")) 
		end if 
		if request.form("userpower")="" then 
			userpower=20 
		else 
			userpower=int(request.form("userpower")) 
		end if 
		if request.form("userart")="" then 
			userart=20 
		else 
			userart=int(request.form("userart")) 
		end if 
		if request.form("vipdate")="" then 
			vipdate=90 
		else 
			vipdate=int(request.form("vipdate")) 
		end if 
		if request.form("userwealth2")="" then 
			money2=20
		else 
			money2=int(request.form("userwealth2")) 
		end if 
		if request.form("userep2")="" then 
			userep2=20 
		else 
			userep2=int(request.form("userep2")) 
		end if 
		if request.form("usercp2")="" then 
			usercp2=20 
		else 
			usercp2=int(request.form("usercp2")) 
		end if 
		if request.form("userpower2")="" then 
			userpower2=20 
		else 
			userpower2=int(request.form("userpower2")) 
		end if 
		if request.form("userart2")="" then 
			userart2=20 
		else 
			userart2=int(request.form("userart2")) 
		end if 
		if request.form("notifydays")="" then 
			notidydays=5 
		else 
			notifydays=int(request.form("notifydays")) 
		end if 
	end if 
	set rs= server.createobject ("adodb.recordset") 
	sql = "select * from vip" 
	rs.Open sql,conn,1,3 
	if not(rs.eof and rs.bof) then 
		rs("wealth") = money 
		rs("userep") = userep 
		rs("usercp") = usercp 
		rs("power") = userpower 
		rs("article") = userart 
		rs("vipdate")=vipdate           
		rs("wealth2") = money2 
		rs("userep2") = userep2 
		rs("usercp2") = usercp2 
		rs("power2") = userpower2 
		rs("article2") = userart2 
		rs("notifydays") = notifydays 
		rs.Update 
	end if 
	rs.close 
	set rs=nothing 
	response.write "�޸ĳɹ�!" 
end sub 

sub addman() 
	dim rs1,sql1,rs,sql,incept,vipdate
	dim msgcontent
	if request("vipname")="" then	
		response.write "��������дVIP�����˰�"
		exit sub
	else
		incept=CheckStr(request("vipname"))
		incept=split(incept,",")
	end if
	sql="select * from [vip]"     
	set rs1=conn.execute(sql)
	vipdate=rs1("vipdate")
	rs1.close
	for i=0 to ubound(incept)
		sql="select vip from [user] where username='"&replace(incept(i),"'","")&"'"
		set rs=conn.execute(sql)
		if rs.eof and rs.bof then
			response.write "��̳û��"&replace(incept(i),"'","")&"����û����������VIP����д�����<BR>"
		else
			if isnull(rs(0)) or rs(0)=4 or rs(0)=0 then
				conn.execute("update [user] set vip=1,vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where username='"&replace(incept(i),"'","")&"'")
				msgcontent="����VIP�ʸ��Ѿ���Ч��лл����"& forum_info(0) &"��֧�֣�"&CHR(10)&"���Ļ�Ա��Ч��Ϊ"&vipdate&"��,��������Ч���ںú�ʹ��VIP��Ա��Ȩ����"
				conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&forum_info(0)&"','������ʽ��ΪVIP��Ա��','"&msgcontent&"',Now(),0,1)")
				response.write "��ӳɹ���<BR>"
			end if
		end if
		rs.close
	next
end sub

sub cleanvip()
	dim vipdate,msgcontent,vipmod,money2,userep2,usercp2,power2,article2
	set rs=conn.execute("select vipdate,vipmod,wealth2,userep2,usercp2,power2,article2 from [vip]")
	vipdate=rs(0)
	vipmod=rs(1)
	money2=rs(2)
	userep2=rs(3)
	usercp2=rs(4)
	power2=rs(5)
	article2=rs(6)
	rs.close
	set rs=conn.execute("select username,vipdate,userwealth,userep,usercp,userpower,article,userwealth2,userep2,usercp2,userpower2,article2 from [User] where vip=1 or vip=3")
	do while not rs.eof
		if datediff("d",rs(1),now)>vipdate then
			if vipmod<>1 then
				if (rs("userwealth")-rs("userwealth2")>=money2 or money2<=0) and (rs("userep")-rs("userep2")>=userep2 or userep2<=0) and (rs("usercp")-rs("usercp2")>=usercp2 or usercp2<=0) and (rs("userpower")-rs("userpower2")>=power2 or power2<=0) and (rs("article")-rs("article2")>=article2 or article2<=0) then
					conn.execute ("update [user] set vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where username='"&rs(0)&"'")
					msgcontent="����VIP�ʸ��ڣ�ϵͳ�Ѿ������Զ���Ϊ����VIP��Ա��лл����"& forum_info(0) &"��֧�֣�"
					conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&rs(0)&"','"&forum_info(0)&"','����VIP��Ա�ʸ����ڳɹ���','"&msgcontent&"',Now(),0,1)")
				else
					conn.execute ("update [User] set vip=0 where username='"&rs(0)&"'")
					msgcontent="����VIP�ʸ��Ѿ����ڣ���������δ�����ٴ������������лл����"& forum_info(0) &"��֧�֣�"&CHR(10)&"�����Ը�⣬������������������������VIP��Ա��"
					conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&rs(0)&"','"&forum_info(0)&"','����VIP��Ա�ʸ�����ʧ�ܣ�','"&msgcontent&"',Now(),0,1)")
				end if
			else
				conn.execute ("update [User] set vip=0,vipdate=now() where username='"&rs(0)&"'")
				msgcontent="����VIP�ʸ��Ѿ����ڣ�лл����"& forum_info(0) &"��֧�֣�"&CHR(10)&"�����Ը�⣬�������ٴ�����VIP��Ա��"
				conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&rs(0)&"','"&forum_info(0)&"','����VIP��Ա�ʸ��Ѿ����ڣ�','"&msgcontent&"',Now(),0,1)")
			end if
		end if
		rs.movenext
	loop
	response.write "���ڹ����û��ɹ�"
end sub%> 
</html>