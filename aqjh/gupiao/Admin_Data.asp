<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-���п������</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<table  width="98%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>���й�������</b></font></td>
	</tr>
	<tr><td bgcolor="#E4E8EF"> 
<% 
call AdminHead()
if not master then
	errmess="<li>��ҳΪ����Աר�ã���û�й���ҳ��Ȩ�ޣ�"
	call endinfo(1)
else 
	select case request("action")
		case "Update1"		
			call Update1()
		case "Update2"		
			call Update2()
		case "Update3"		
			call Update3()			
		case "ExecuteSql"		
			call ExecuteSql()
		case "ClearEvent"
			 call ClearEvent()
		case else
			call main()
	end select
end if
%>
<br>
</td></tr>
<tr><td height=32 background="images/footer.gif" valign=middle></td></tr>
</table>
</body></html>
<%
CloseDatabase		'�ر����ݿ�
'=====================================
sub main()
%>
<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
	<tr>
		<td align="center" colspan="2" height="25" background="Images/Header.gif"><b>���ݿ����</b></td>
	</tr>
	<tr>
		<th colspan="2" height="25" class="admin" align="left">�������ݹ���</th>
	</tr>
	<form method="POST" action="?action=Update1" name="f1">	
	<tr>
		<td class="forumRow" width="50%"><b>���¼����Ʊ��ӪȨ</b><br>���ݹ���ֹ������¼����Ʊ�ľ�Ӫ��<br>����ÿ��һ��ʱ������һ��</td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="���¾�ӪȨ">
		</td>
	</tr>
	</form>
	<form method="POST" action="?action=Update2" name="f2">
	<tr>
		<td class="forumRow" width="50%"><b>���¼����ʲ�����</b><br>���¼����û��ֹ��������ֹ���ֵ�����ʲ�ֵ<br>����ÿ��һ��ʱ������һ��</td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="ˢ���ʲ�����">
		</td>
	</tr>
	</form>
	<form method="POST" action="?action=Update3" name="f3">
	<tr>
		<td class="forumRow" width="50%"><b>���¹�Ʊ����</b><br></td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="���¹�Ʊ����">
		</td>
	</tr>
	</form>	
	<tr>
		<th colspan="2" height="25" class="admin" align="left">����¼�����</th>
	</tr>	
	<form method="POST" action="?action=ClearEvent" name="f4">
	<tr>
		<td class="forumRow" width="50%"><b>�������¼���¼</b><br>ɾ�����е�����¼���¼</td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="ִ��">
		</td>
	</tr>
	</form>
	<form method="POST" action="?action=ClearEvent&reaction=created" name="f4">
	<tr>
		<td class="forumRow" width="50%"><b>�ؽ�����¼���</b><br>ɾ�����е�����¼���¼�����½�������¼���[RndEvent]</td> 
		<td class="forumRow" width="50%">
		<input type=submit name="submit" value="ִ��">
		</td>
	</tr>
	</form>		
	<tr>
		<th colspan="2" height="25" class="admin" align="left">SOL���ִ�в���</th>
	</tr>	
	<form method="POST" action="?action=ExecuteSql" name="f3">
	<tr>
		<td class="forumRow" colspan="2">
		<br>&nbsp;&nbsp;&nbsp;&nbsp;���������޸߼�����SQL��̱Ƚ���Ϥ���û���������ֱ������sqlִ����䣬����delete from [�ͻ�] where username='test'����ɾ���û����ڲ���ǰ�����ؿ�������ִ������Ƿ���ȷ��������������ֱ�ӶԹ�Ʊ���ݿ���в�����ִ�к󲻿ɻָ���<br><br>
 		������SQL��䣺
		<input type=text name="sql" size=100>
		<input type=submit name=submit value="ִ��">
		<br><br>		
		</td>
	</tr>
	</form>		
</table>	
<% 
end sub
'----------���¼����Ʊ��ӪȨ-------------
sub Update1()
	dim rst
    set rs=conn.execute("select sid from [��Ʊ] order by sid")
	do while not rs.eof
		set rst=conn.execute("select top 1 �ʺ�,Uid from [��] where sid="&rs(0)&" and Uid>0 order by �ֹ��� desc")
		if rst.eof then	
			conn.execute("update [��Ʊ] set ��Ӫ��=null,Uid=0 where sid="&rs(0))
		else
			conn.execute("update [��Ʊ] set ��Ӫ��='"&rst(0)&"',Uid="&rst(1)&" where sid="&rs(0))
		end if	
		rs.movenext
	loop
	rst.close
	rs.close	
	sucmess="<li>���ݿ���³ɹ�"
	call endinfo(1)
end sub
'-----------���¼����ʲ�����-------------
sub Update2()
	dim rst,TotalFund,StockTypeNum
	set rs=conn.execute("select id from [�ͻ�] order by id")
	do while not rs.eof 
		TotalFund=0 :	StockTypeNum=0	
		sql="select d.�ֹ���,g.��ǰ�۸� from [��] d inner join [��Ʊ] g on d.sid=g.sid where d.uid="&rs(0)
		set rst=conn.execute(sql)
		do while not rst.eof
			TotalFund=TotalFund+rst(0)*rst(1)
			StockTypeNum=StockTypeNum+1
			rst.movenext
		loop
		conn.execute("update [�ͻ�] set ���ʽ�=�ʽ�+"&TotalFund&",�ֹ�����="&StockTypeNum&" where id="&rs(0))
		rs.movenext
	loop
	rst.close
	rs.close
	sucmess="<li>�û��ֹ��������ֹ���ֵ�����ʲ�ֵ���³ɹ�"
	call endinfo(1)
end sub

sub Update3()
	dim StockNum,TongJi
	set rs=conn.execute("select sid,TongJi from [��Ʊ] order by sid")
	do while not rs.eof 
		TongJi=rs(1)
		StockNum=conn.execute("select sum(�ֹ���) from [��] where sid="&rs(0))(0)
		if StockNum="" or isnull(StockNum) then StockNum=0
		if TongJi="" or isnull(TongJi) then TongJi="0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
		conn.execute("update [��Ʊ] set ʣ��ɷ�=�ܹɷ�-"&StockNum&",TongJi='"&TongJi&"' where sid="&rs(0))
		rs.movenext
	loop
	rs.close
	sucmess="<li>���³ɹ�"
	call endinfo(1)
end sub
'------SQL���ִ�в���-----------
sub ExecuteSql()
	if request("sql")="" then
		errmess=errmess+"<br>"+"<li>������SQL���" 
		call endinfo(1)		
		exit sub
	else
		sql=request("sql")
	end if	
	
	On Error Resume Next 
	conn.Execute(sql)
	if err.number="0" then
		sucmess=sucmess+"<br>"+"<li>�����SQL��䣺"&sql
		sucmess=sucmess+"<br>"+"<li>���ִ�гɹ�"
	 	call endinfo(1)
	else
		errmess=errmess+"<br>"+"<li>�����SQL��䣺"&sql
		errmess=errmess+"<br>"+"<li>��������⣬����������£�"
		errmess=errmess+"<br><li>"&Err.Description
		err.clear
		call endinfo(1)
	end if		
	
end sub

sub ClearEvent()
	if request("reaction")="created" then
		'ɾ���¼���[�¼�]
		conn.execute("drop table [RndEvent]")	
		'�ؽ��¼��� RndEvent
		sql="CREATE TABLE [RndEvent] (ID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,content varchar(255),AddTime datetime default now())"
		conn.execute(sql)
		sucmess=sucmess+"<br>"+"<li>�ؽ��¼���ɹ�"
	else
		'ɾ����������¼�
		conn.execute("delete from [RndEvent]")
		sucmess=sucmess+"<br>"+"<li>��������¼��Ѿ����" 
	end if
	call endinfo(1)	
end sub
%>