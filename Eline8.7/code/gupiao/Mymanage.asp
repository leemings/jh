<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="�����ֽ�������Ʊ�г�">
<title><%=Gupiao_Setting(5)%>-�����ʻ�����</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<%
	dim User_Setting
	User_Setting=conn.execute("select top 1 User_Setting from GupiaoConfig order by id")(0)
	User_Setting=split(User_Setting,"|")
	
	select case request("action")
		case "savecun"
			call savecun()
		case "savequ"
			call savequ()
		case "new"
			call Create()	'����	
		case else
			call main()
	end select		
	CloseDatabase		'�ر����ݿ�
%>
</body>
</html>
<%	
'---------�ʻ�����----------
sub main() 
	sql="select * from �ͻ� where id="&MyUserID
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		call Create()	'����
	else	
%>
<table  width="420" height=300 border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9></TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 class=big background="images/header.gif">��Ʊ�ʻ�����</td>
	</tr>
	<tr>
		<td valign=center BGCOLOR="#E4E4EF">

      	<p align=center><font color="#3366CC">���Ĺ�Ʊ�ʻ��������ʽ� <font color=red><%=int(rs("�ʽ�"))%></font> Ԫ</font></p>
<%
		'sqlbbs="select userWealth from [user] where username='"&membername&"'"
		sqlbbs="select ���� as userWealth from [�û�] where ����='"&membername&"'"
		set rsbbs=connbbs.execute(sqlbbs)
      	response.Write "<p align=center><font color=""#000099"">�����������ֽ� <font color=red>"&rsbbs("userWealth")&"</font> Ԫ</font></p>"
		rsbbs.close
%>
      	<p align=center style="color:#000000"> 
			<form method=POST action="?action=savequ" name=form1>
			�ӹ�Ʊ�ʻ���<input type=text name=money  size=10>		<input type=submit value=ȡ�� name=submit1>
			</form>
		</p>
      	<p align=center> 
			<form method=POST action="?action=savecun" name=form2>
			����Ʊ�ʻ���<input type=text name=money  size=10>		<input type=submit value=��� name=submit2>
			</form>
		</p>			
		</td>
	</tr>
	<TR><TD height=30 background="images/footer.gif" align="center"><A href="stock.asp"><font color="#0000ff">���ع�Ʊ���״���</font></A>
	</TD></TR>
</table>
<%
	end if	
end sub 
'---------�ʻ���Ǯ----------
sub savequ()
	dim  money
	if clng(User_Setting(0))=0 then
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
	elseif money>csng(User_Setting(2)) and csng(User_Setting(2))>0 then
		errmess="<li>��Ĺ�Ʊ�����˴���˵��ÿ�����ܳ��� "&User_Setting(2)& " Ԫ"
		call endinfo(2)	
	else
		sql="update �ͻ� set �ʽ�=�ʽ�-"&money&",���ʽ�=���ʽ�-"&money&" where id="&MyUserID
		conn.execute sql
		'sqlbbs="update [user] set userWealth=userWealth+" & money & " where username='"&membername&"'"
		sqlbbs="update [�û�] set ����=����+" & money & " where ����='"&membername&"'"
		connbbs.execute sqlbbs
		sucmess="<li>��ӹ�Ʊ�ʻ�ȡ���� "&money&" �����ӣ��ú��˰�"
		call endinfo(1)
	end if
end sub
'-----------------�ʻ���Ǯ--------------------
sub savecun()
	dim money,bbsCash
	if clng(User_Setting(1))=0 then
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
	if money>csng(User_Setting(3)) and csng(User_Setting(3))>0 then
		errmess="<li>��Ǯ����ÿ�����ֻ�ܴ��� "&User_Setting(3)& " Ԫ"
		call endinfo(2)
		exit sub
	end if	
	'sqlbbs="select * from [user] where username='"& membername&"'"
	sqlbbs="select ���� as userWealth from [�û�] where ����='"& membername&"'"
	set rsbbs=connbbs.execute(sqlbbs)
	bbsCash=rsbbs("userWealth")
	if money>bbsCash then 
		errmess="<li>��Ĺ�Ʊ�����˴���˵����������ô��Ǯ���������"
		call endinfo(2)
	else
		'sqlbbs="update [user] set userWealth=userWealth-"& money & " where username='" & membername & "'"
		sqlbbs="update [�û�] set ����=����-"& money & " where ����='" & membername & "'"
		connbbs.execute sqlbbs
		sql="update �ͻ� set �ʽ�=�ʽ�+"&money&",���ʽ�=���ʽ�+"&money&" where id="&MyUserID
		conn.execute sql
		sucmess="<li>���Ѿ��� "&money&" �����Ӵ������Ĺ�Ʊ�ʻ�"
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
	KaiHu_Setting=conn.execute("select top 1 KaiHu_Setting from GupiaoConfig order by id")(0)
	KaiHu_Setting=split(KaiHu_Setting,"|")
	if clng(KaiHu_Setting(0))=0 then
		errmess="<li>������ͣ�������������Ա��ϵ"
		call endinfo(1)
		exit sub
	end if

	sqlbbs="select ����,���� from [�û�] where ����='"&membername&"'"
	set rsbbs=connbbs.execute(sqlbbs)	

'	if rsbbs(0)<clng(KaiHu_Setting(1)) then
'		errmess=errmess+"<li>�������������˿������ٷ�����Ϊ"&KaiHu_Setting(1)&"�������ڲ����ϱ�Ҫ����ʱ���ܿ���"
'		founderr=true
'	end if
	if rsbbs("����")<clng(KaiHu_Setting(2)) then
		errmess=errmess+"<li>�������������˿�����������ֵΪ"&KaiHu_Setting(2)&"���������ϱ�Ҫ����ʱ���ܿ���"
		founderr=true
	end if
'	if rsbbs(2)<clng(KaiHu_Setting(3)) then
'		errmess=errmess+"<li>�������������˿������پ���ֵΪ"&KaiHu_Setting(3)&"���������ϱ�Ҫ����ʱ���ܿ���"
'		founderr=true
'	end if
	if rsbbs("����")<clng(KaiHu_Setting(4)) then
		errmess=errmess+"<li>�������������˿��������ֽ�ֵΪ"&KaiHu_Setting(4)&"���������ϱ�Ҫ����ʱ���ܿ���"
		founderr=true
	end if
	rsbbs.close
	set rsbbs=nothing
					
	if founderr then
		call endinfo(1)
	else		
		sql="insert into �ͻ�(�ʺ�,�ʽ�,���ʽ�) values('" & membername & "',"&KaiHu_Setting(6)&","&KaiHu_Setting(6)&")"
		conn.execute sql
		sucmess="<li>��Ʊ�ʻ�����ɹ���<br><li><font color=#000099>�������Ǹ����񣬹�Ʊ�г��㷢�ȱ�����<font color=#990000><b> "&KaiHu_Setting(6)&" </b></font> �����ӽ�������ʻ���ְɣ�</font> "
		call endinfo(2)
	end if	
end sub
%>