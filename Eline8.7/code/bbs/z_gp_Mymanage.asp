<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="�����˺Ź���"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
dim User_Setting
User_Setting=gp_conn.execute("select top 1 User_Setting from GupiaoConfig order by id")(0)
User_Setting=split(User_Setting,"|")
User_Setting(0)=CLNG(User_Setting(0))
User_Setting(1)=CLNG(User_Setting(1))
User_Setting(2)=CLNG(User_Setting(2))
User_Setting(3)=CLNG(User_Setting(3))%>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
<%select case request("action")
case "savecun"
	call savecun()
case "savequ"
	call savequ()
case "new"
	call Create()	'����	
case else
	call main()
end select%>
</table>
<%call activeonline()
call footer()
'---------�ʻ�����----------
sub main() 
	sql="select * from [KeHu] where id="&MyUserID&" and suoding<2"
	set rs=gp_conn.execute(sql)
	if rs.eof and rs.bof then
		call Create()	'����
	else%>
		<tr>
			<th valign=center align=middle height=25>��Ʊ�ʻ�����</th>
		</tr>
		<tr>
			<td valign=center class=tablebody1><p align=center><font color="#3366CC">���Ĺ�Ʊ�ʻ��������ʽ� <font color=red><%=int(rs("ZiJin"))%></font> Ԫ</font></p><%
				sqlbbs="select userWealth from [user] where username='"&membername&"'"
				set rsbbs=conn.execute(sqlbbs)
      	response.Write "<p align=center><font color=""#000099"">�����������ֽ� <font color=red>"&rsbbs("userWealth")&"</font> Ԫ</font></p>"
				rsbbs.close
				%><p align=center><form method=POST action="?action=savequ" name=form1>�ӹ�Ʊ�ʻ���<input type=text name=money  size=10>		<input type=submit value=ȡ�� name=submit1></form></p>
      	<p align=center><form method=POST action="?action=savecun" name=form2>����Ʊ�ʻ���<input type=text name=money  size=10>		<input type=submit value=��� name=submit2></form></p>			
			</td>
		</tr>
		<TR>
			<Td class=tablebody2 height=25  align="center"><A href="z_gp_Gupiao.asp"><b>���ع�Ʊ���״���</b></A></Td>
		</TR>
	<%end if	
end sub 
'---------�ʻ���Ǯ----------
sub savequ()
	dim  money
	if User_Setting(0)=0 then
		errmess="<li>�����ʻ���ʱ������Ǯ����"
		call endinfo(2)
		exit sub	
	end if 
	if not isnumeric(request.form("money")) then
		errmess="<li>��û������������������Ĳ�������"
		call endinfo(2)
		exit sub
	elseif request.form("money")<=0 then
		errmess="<li>�������������С�ڵ�����"
		call endinfo(2)
		exit sub		
	end if
	money=int(request.form("money"))
	
	if money>MyCash then 
		errmess="<li>��Ĺ�Ʊ�����˴���˵����û�а���׬��ô��Ǯѽ��"
		call endinfo(2)
	elseif money>User_Setting(2) and User_Setting(2)>0 then
		errmess="<li>��Ĺ�Ʊ�����˴���˵��ÿ�����ܳ��� "&User_Setting(2)& " Ԫ"
		call endinfo(2)	
	else
		sql="update [KeHu] set ZiJin=ZiJin-"& money&",ZongZiJin=ZongZiJin-"&money&" where id="&MyUserID
		gp_conn.execute sql
		sqlbbs="update [user] set userWealth=userWealth+" & money & " where username='"&membername&"'"
		conn.execute sqlbbs
		sucmess="<li>��ӹ�Ʊ�ʻ�ȡ���� "&money&" Ԫ���ú��˰�"
		call endinfo(1)
	end if
end sub
'-----------------�ʻ���Ǯ--------------------
sub savecun()
	dim money,bbsCash
	if User_Setting(1)=0 then
		errmess="<li>�����ʻ���ʱ���ܽ��д�Ǯ����"
		call endinfo(2)
		exit sub	
	end if 	
	if not isnumeric(request.form("money")) then
		errmess="<li>��û������������������Ĳ�������"
		call endinfo(2)
		exit sub
	elseif request.form("money")<=0 then
		errmess="<li>�������������С�ڵ�����"
		call endinfo(2)
		exit sub		
	end if
	money=int(request.form("money"))
	if money>User_Setting(3) and User_Setting(3)>0 then
		errmess="<li>��Ǯ����ÿ�����ֻ�ܴ��� "&User_Setting(3)& " Ԫ"
		call endinfo(2)
		exit sub
	end if	
	sqlbbs="select * from [user] where username='"& membername&"'"
	set rsbbs=conn.execute(sqlbbs)
	bbsCash=rsbbs("userWealth")
	if money>bbsCash then 
		errmess="<li>��Ĺ�Ʊ�����˴���˵����������ô��Ǯ���������"
		call endinfo(2)
	else
		sqlbbs="update [user] set userWealth=userWealth-"& money & " where username='" & membername & "'"
		conn.execute sqlbbs
		sql="update [KeHu] set ZiJin=ZiJin+" & money&",ZongZiJin=ZongZiJin+"&money&" where id="&MyUserID
		gp_conn.execute sql
		sucmess="<li>���Ѿ��� "&money&" Ԫ�������Ĺ�Ʊ�ʻ�"
		call endinfo(1)
	end if
end sub
'-----------------����-----------------
sub Create()
	if HaveAccount then
		errmess="<li>���Ѿ��й�Ʊ�ʻ��ˣ���ô����������"
		call endinfo(1)
		exit sub
	end if
			
	dim KaiHu_Setting
	KaiHu_Setting=gp_conn.execute("select top 1 KaiHu_Setting from GupiaoConfig order by id")(0)
	KaiHu_Setting=split(KaiHu_Setting,"|")
	KaiHu_Setting(0)=CLNG(KaiHu_Setting(0))
	KaiHu_Setting(1)=CLNG(KaiHu_Setting(1))
	KaiHu_Setting(2)=CLNG(KaiHu_Setting(2))
	KaiHu_Setting(3)=CLNG(KaiHu_Setting(3))
	KaiHu_Setting(4)=CLNG(KaiHu_Setting(4))
	KaiHu_Setting(6)=CLNG(KaiHu_Setting(6))
	if KaiHu_Setting(0)=0 then
		errmess="<li>������ͣ�������������Ա��ϵ"
		call endinfo(1)
		exit sub
	end if

	sqlbbs="select Article,userCP,userEP,userWealth from [user] where username='"&membername&"'"
	set rsbbs=conn.execute(sqlbbs)	

	if rsbbs(0)<KaiHu_Setting(1) then
		errmess=errmess+"<li>�������������˿������ٷ�����Ϊ"&KaiHu_Setting(1)&"�������ڲ����ϱ�Ҫ����ʱ���ܿ���"
		founderr=true
	end if
	if rsbbs(1)<KaiHu_Setting(2) then
		errmess=errmess+"<li>�������������˿�����������ֵΪ"&KaiHu_Setting(2)&"���������ϱ�Ҫ����ʱ���ܿ���"
		founderr=true
	end if
	if rsbbs(2)<KaiHu_Setting(3) then
		errmess=errmess+"<li>�������������˿������پ���ֵΪ"&KaiHu_Setting(3)&"���������ϱ�Ҫ����ʱ���ܿ���"
		founderr=true
	end if
	if rsbbs(3)<KaiHu_Setting(4) then
		errmess=errmess+"<li>�������������˿��������ֽ�ֵΪ"&KaiHu_Setting(4)&"���������ϱ�Ҫ����ʱ���ܿ���"
		founderr=true
	end if
	rsbbs.close
	set rsbbs=nothing
					
	if founderr then
		call endinfo(1)
	else		
		if MyUserID<>0 then
			sql="update [KeHu] set SuoDing=0,KaiHuRiQi=Now(),ZuiHouRiQi=Now() where id="&myuserid
			gp_conn.execute sql
			sucmess="<li>��Ʊ�ʻ�����ɹ���<br><li><font color=#000099>��������ǰ�Ѿ���ȡ�������ʽ𣬱��β��ٷ��ţ�</font> "
		else
			sql="insert into [KeHu](ZhangHao,ZiJin,ZongZiJin) values('" & membername & "',"&KaiHu_Setting(6)&","&KaiHu_Setting(6)&")"
			gp_conn.execute sql
			sucmess="<li>��Ʊ�ʻ�����ɹ���<br><li><font color=#000099>�������Ǹ����񣬹�Ʊ�г��㷢�ȱ�����<font color=#990000><b> "&KaiHu_Setting(6)&" </b></font> Ԫ��������ʻ���ְɣ�</font> "
		end if
		call endinfo(2)
	end if	
end sub
%>