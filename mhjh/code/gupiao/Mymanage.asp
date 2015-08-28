<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="魔幻江湖股票市场">
<title><%=Gupiao_Setting(5)%>-个人帐户管理</title>
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
			call Create()	'开户	
		case else
			call main()
	end select		
	CloseDatabase		'关闭数据库
%>
</body>
</html>
<%	
'---------帐户管理----------
sub main() 
	sql="select * from 客户 where id="&MyUserID
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		call Create()	'开户
	else	
%>
<table  width="420" height=300 border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9></TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 class=big background="images/header.gif">股票帐户管理</td>
	</tr>
	<tr>
		<td valign=center BGCOLOR="#E4E4EF">

      	<p align=center><font color="#3366CC">您的股票帐户里现有资金 <font color=red><%=int(rs("资金"))%></font> 元</font></p>
<%
		'sqlbbs="select userWealth from [user] where username='"&membername&"'"
		sqlbbs="select 银两 as userWealth from 用户 where 姓名='"&membername&"'"
		set rsbbs=connbbs.execute(sqlbbs)
      	response.Write "<p align=center><font color=""#000099"">您的身上有现金 <font color=red>"&rsbbs("userWealth")&"</font> 元</font></p>"
		rsbbs.close
%>
      	<p align=center style="color:#000000"> 
			<form method=POST action="?action=savequ" name=form1>
			从股票帐户提款：<input type=text name=money  size=10>		<input type=submit value=取款 name=submit1>
			</form>
		</p>
      	<p align=center> 
			<form method=POST action="?action=savecun" name=form2>
			存款到股票帐户：<input type=text name=money  size=10>		<input type=submit value=存款 name=submit2>
			</form>
		</p>			
		</td>
	</tr>
	<TR><TD height=30 background="images/footer.gif" align="center"><A href="stock.asp"><font color="#0000ff">返回股票交易大厅</font></A>
	</TD></TR>
</table>
<%
	end if	
end sub 
'---------帐户提钱----------
sub savequ()
	dim  money
	if clng(User_Setting(0))=0 then
		errmess="<li>股市帐户暂时不能提钱出来"
		call endinfo(2)
		exit sub	
	end if 
	if not isnumeric(request.form("money")) then
		errmess="<li>您没有输入提款金额或者输入的不是数字"
		call endinfo(2)
		exit sub
	elseif request.form("money")<=0 then
		errmess="<li>您输入的提款金额不能小于等于零"
		call endinfo(2)
		exit sub		
	end if
	money=int(request.form("money"))
	
	if money>MyCash then 
		errmess="<li>你的股票经纪人大声说，我没有帮你赚这么多钱呀？"
		call endinfo(2)
	elseif money>csng(User_Setting(2)) and csng(User_Setting(2))>0 then
		errmess="<li>你的股票经纪人大声说，每次提款不能超过 "&User_Setting(2)& " 元"
		call endinfo(2)	
	else
		sql="update 客户 set 资金=资金-"&money&",总资金=总资金-"&money&" where id="&MyUserID
		conn.execute sql
		'sqlbbs="update [user] set userWealth=userWealth+" & money & " where username='"&membername&"'"
		sqlbbs="update 用户 set 银两=银两+" & money & " where 姓名='"&membername&"'"
		connbbs.execute sqlbbs
		sucmess="<li>你从股票帐户取出了 "&money&" 两银子，拿好了啊"
		call endinfo(1)
	end if
end sub
'-----------------帐户存钱--------------------
sub savecun()
	dim money,bbsCash
	if clng(User_Setting(1))=0 then
		errmess="<li>股市帐户暂时不能进行存钱操作"
		call endinfo(2)
		exit sub	
	end if 	
	if not isnumeric(request.form("money")) then
		errmess="<li>您没有输入提款金额或者输入的不是数字"
		call endinfo(2)
		exit sub
	elseif request.form("money")<=0 then
		errmess="<li>您输入的提款金额不能小于等于零"
		call endinfo(2)
		exit sub		
	end if
	money=int(request.form("money"))
	if money>csng(User_Setting(3)) and csng(User_Setting(3))>0 then
		errmess="<li>存钱操作每次最多只能存入 "&User_Setting(3)& " 元"
		call endinfo(2)
		exit sub
	end if	
	'sqlbbs="select * from [user] where username='"& membername&"'"
	sqlbbs="select 银两 as userWealth from 用户 where 姓名='"& membername&"'"
	set rsbbs=connbbs.execute(sqlbbs)
	bbsCash=rsbbs("userWealth")
	if money>bbsCash then 
		errmess="<li>你的股票经纪人大声说，你哪里这么多钱，想坑我吗？"
		call endinfo(2)
	else
		'sqlbbs="update [user] set userWealth=userWealth-"& money & " where username='" & membername & "'"
		sqlbbs="update 用户 set 银两=银两-"& money & " where 姓名='" & membername & "'"
		connbbs.execute sqlbbs
		sql="update 客户 set 资金=资金+"&money&",总资金=总资金+"&money&" where id="&MyUserID
		conn.execute sql
		sucmess="<li>你已经把 "&money&" 两银子存进了你的股票帐户"
		call endinfo(1)
	end if
end sub
'-----------------开户-----------------
sub Create()
	if HaveAccount then
		errmess="<li>您已经有股票帐户了，怎么还来开户？"
		call endinfo(1)
		exit sub
	end if
			
	dim KaiHu_Setting
	KaiHu_Setting=conn.execute("select top 1 KaiHu_Setting from GupiaoConfig order by id")(0)
	KaiHu_Setting=split(KaiHu_Setting,"|")
	if clng(KaiHu_Setting(0))=0 then
		errmess="<li>股市暂停开户，请与管理员联系"
		call endinfo(1)
		exit sub
	end if

	sqlbbs="select 道德,银两,积分 from 用户 where 姓名='"&membername&"'"
	set rsbbs=connbbs.execute(sqlbbs)	

	if rsbbs("道德")<clng(KaiHu_Setting(2)) then
		errmess=errmess+"<li>本股市已限制了开户最少道德值为"&KaiHu_Setting(2)&"，您不符合本要求，暂时不能开户"
		founderr=true
	end if
'	if rsbbs("积分")<clng(KaiHu_Setting(3)) then
'		errmess=errmess+"<li>本股市已限制了开户最少积分为"&KaiHu_Setting(3)&"，您不符合本要求，暂时不能开户"
'		founderr=true
'	end if
	if rsbbs("银两")<clng(KaiHu_Setting(4)) then
		errmess=errmess+"<li>本股市已限制了开户最少现金值为"&KaiHu_Setting(4)&"，您不符合本要求，暂时不能开户"
		founderr=true
	end if
	rsbbs.close
	set rsbbs=nothing
					
	if founderr then
		call endinfo(1)
	else		
		sql="insert into 客户(帐号,资金,总资金) values('" & membername & "',"&KaiHu_Setting(6)&","&KaiHu_Setting(6)&")"
		conn.execute sql
		sucmess="<li>股票帐户申请成功！<br><li><font color=#000099>看到你是个菜鸟，股票市场便发慈悲送了<font color=#990000><b> "&KaiHu_Setting(6)&" </b></font> 两银子进了你的帐户里，乐吧？</font> "
		call endinfo(2)
	end if	
end sub
%>