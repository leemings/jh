<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="���ݿ����"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
if not master then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
else
	call AdminHead()%>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
	<%select case request("action")
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
	call footer()%>
	</table>
<%end if
'=====================================
sub main()%>
  <tr> 
    <th align="center" colspan="2" height="25">���ݿ����</th>
  </tr>
  <tr> 
    <td class=tablebody2 colspan="2" height="25" align="left"><b>�������ݹ���</b></td>
	</tr>
	<form method="POST" action="?action=Update1" name="f1">	
	<tr>
    <td class=tablebody1 width="50%"><b>���¼����Ʊ��ӪȨ</b><br>���ݹ���ֹ������¼����Ʊ�ľ�Ӫ��<br>����ÿ��һ��ʱ������һ��</td>
    <td class=tablebody1 width="50%"> <input type=submit name="submit" value="���¾�ӪȨ"></td>
	</tr>
	</form>
	<form method="POST" action="?action=Update2" name="f2">
	<tr>
    <td class=tablebody1 width="50%"><b>���¼����ʲ�����</b><br>���¼����û��ֹ��������ֹ���ֵ�����ʲ�ֵ<br>����ÿ��һ��ʱ������һ��</td>
    <td class=tablebody1 width="50%"> <input type=submit name="submit" value="ˢ���ʲ�����"></td>
	</tr>
	</form>
	<form method="POST" action="?action=Update3" name="f3">
	<tr>
    <td class=tablebody1 width="50%"><b>���¹�Ʊ����</b><br></td>
    <td class=tablebody1 width="50%"> <input type=submit name="submit" value="���¹�Ʊ����"></td>
	</tr>
	</form>	
	<tr>
    <td class=tablebody2 colspan="2" height="25" align="left"><B>����¼�����</b></td>
	</tr>	
	<form method="POST" action="?action=ClearEvent" name="f4">
	<tr>
    <td class=tablebody1 width="50%"><b>�������¼���¼</b><br>ɾ�����е�����¼���¼</td>
    <td class=tablebody1 width="50%"> <input type=submit name="submit" value="ִ��"></td>
	</tr>
	</form>
	<form method="POST" action="?action=ClearEvent&reaction=created" name="f4">
	<tr>
		<td class="tablebody1" width="50%"><b>�ؽ�����¼���</b><br>ɾ�����е�����¼���¼�����½�������¼���[RndEvent]</td> 
		<td class="tablebody1" width="50%"> <input type=submit name="submit" value="ִ��"></td>
	</tr>
	</form>		
	<tr>
		<td class=tablebody2 colspan="2" height="25" align="left"><b>SOL���ִ�в���</b></td>
	</tr>	
	<form method="POST" action="?action=ExecuteSql" name="f3">
	<tr>
		<td class=tablebody1 colspan="2" align=center><br>&nbsp;&nbsp;&nbsp;&nbsp;���������޸߼�����SQL��̱Ƚ���Ϥ���û���������ֱ������sqlִ����䣬����delete from [KeHu] where username='test'����ɾ���û����ڲ���ǰ�����ؿ�������ִ������Ƿ���ȷ��������������ֱ�ӶԹ�Ʊ���ݿ���в�����ִ�к󲻿ɻָ���<br><br>������SQL��䣺<input type=text name="sql" size=100> <input type=submit name=submit value="ִ��"><br><br></td>
	</tr>
	</form>		
<%end sub
'----------���¼����Ʊ��ӪȨ-------------
sub Update1()
	dim rst
    set rs=gp_conn.execute("select sid from [GuPiao] order by sid")
	do while not rs.eof
		set rst=gp_conn.execute("select top 1 ZhangHao,Uid from [DaHu] where sid="&rs(0)&" and Uid>0 order by ChiGuShu desc")
		if rst.eof then	
			gp_conn.execute("update [GuPiao] set JingYingZhe=null,Uid=0 where sid="&rs(0))
		else
			gp_conn.execute("update [GuPiao] set JingYingZhe='"&rst(0)&"',Uid="&rst(1)&" where sid="&rs(0))
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
	set rs=gp_conn.execute("select id from [KeHu] where suoding<2 order by id")
	do while not rs.eof 
		TotalFund=0 :	StockTypeNum=0	
		sql="select d.ChiGuShu,g.DangQianJiaGe from [DaHu] d inner join [GuPiao] g on d.sid=g.sid where d.uid="&rs(0)
		set rst=gp_conn.execute(sql)
		do while not rst.eof
			TotalFund=TotalFund+rst(0)*rst(1)
			StockTypeNum=StockTypeNum+1
			rst.movenext
		loop
		gp_conn.execute("update [KeHu] set ZongZiJin=ZiJin+"&TotalFund&",ChiGuZhongLei="&StockTypeNum&" where id="&rs(0))
		rs.movenext
	loop
	rs.close
	sucmess="<li>�û��ֹ��������ֹ���ֵ�����ʲ�ֵ���³ɹ�"
	call endinfo(1)
end sub

sub Update3()
	dim StockNum,TongJi
	set rs=gp_conn.execute("select sid,TongJi from [GuPiao] order by sid")
	do while not rs.eof 
		TongJi=rs(1)
		StockNum=gp_conn.execute("select sum(ChiGuShu) from [DaHu] where sid="&rs(0))(0)
		if StockNum="" or isnull(StockNum) then StockNum=0
		if TongJi="" or isnull(TongJi) then TongJi="0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
		gp_conn.execute("update [GuPiao] set ShengYuGuFen=ZongGuFen-"&StockNum&",TongJi='"&TongJi&"' where sid="&rs(0))
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
	gp_conn.Execute(sql)
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
		gp_conn.execute("drop table [RndEvent]")	
		'�ؽ��¼��� RndEvent
		sql="CREATE TABLE [RndEvent] (ID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,content varchar(255),AddTime datetime default now())"
		gp_conn.execute(sql)
		sucmess=sucmess+"<br>"+"<li>�ؽ��¼���ɹ�"
	else
		'ɾ����������¼�
		gp_conn.execute("delete from [RndEvent]")
		sucmess=sucmess+"<br>"+"<li>��������¼��Ѿ����" 
	end if
	call endinfo(1)	
end sub
%>