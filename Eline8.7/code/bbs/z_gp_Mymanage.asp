<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="个人账号管理"
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
	call Create()	'开户	
case else
	call main()
end select%>
</table>
<%call activeonline()
call footer()
'---------帐户管理----------
sub main() 
	sql="select * from [KeHu] where id="&MyUserID&" and suoding<2"
	set rs=gp_conn.execute(sql)
	if rs.eof and rs.bof then
		call Create()	'开户
	else%>
		<tr>
			<th valign=center align=middle height=25>股票帐户管理</th>
		</tr>
		<tr>
			<td valign=center class=tablebody1><p align=center><font color="#3366CC">您的股票帐户里现有资金 <font color=red><%=int(rs("ZiJin"))%></font> 元</font></p><%
				sqlbbs="select userWealth from [user] where username='"&membername&"'"
				set rsbbs=conn.execute(sqlbbs)
      	response.Write "<p align=center><font color=""#000099"">您的身上有现金 <font color=red>"&rsbbs("userWealth")&"</font> 元</font></p>"
				rsbbs.close
				%><p align=center><form method=POST action="?action=savequ" name=form1>从股票帐户提款：<input type=text name=money  size=10>		<input type=submit value=取款 name=submit1></form></p>
      	<p align=center><form method=POST action="?action=savecun" name=form2>存款到股票帐户：<input type=text name=money  size=10>		<input type=submit value=存款 name=submit2></form></p>			
			</td>
		</tr>
		<TR>
			<Td class=tablebody2 height=25  align="center"><A href="z_gp_Gupiao.asp"><b>返回股票交易大厅</b></A></Td>
		</TR>
	<%end if	
end sub 
'---------帐户提钱----------
sub savequ()
	dim  money
	if User_Setting(0)=0 then
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
	elseif money>User_Setting(2) and User_Setting(2)>0 then
		errmess="<li>你的股票经纪人大声说，每次提款不能超过 "&User_Setting(2)& " 元"
		call endinfo(2)	
	else
		sql="update [KeHu] set ZiJin=ZiJin-"& money&",ZongZiJin=ZongZiJin-"&money&" where id="&MyUserID
		gp_conn.execute sql
		sqlbbs="update [user] set userWealth=userWealth+" & money & " where username='"&membername&"'"
		conn.execute sqlbbs
		sucmess="<li>你从股票帐户取出了 "&money&" 元，拿好了啊"
		call endinfo(1)
	end if
end sub
'-----------------帐户存钱--------------------
sub savecun()
	dim money,bbsCash
	if User_Setting(1)=0 then
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
	if money>User_Setting(3) and User_Setting(3)>0 then
		errmess="<li>存钱操作每次最多只能存入 "&User_Setting(3)& " 元"
		call endinfo(2)
		exit sub
	end if	
	sqlbbs="select * from [user] where username='"& membername&"'"
	set rsbbs=conn.execute(sqlbbs)
	bbsCash=rsbbs("userWealth")
	if money>bbsCash then 
		errmess="<li>你的股票经纪人大声说，你哪里这么多钱，想坑我吗？"
		call endinfo(2)
	else
		sqlbbs="update [user] set userWealth=userWealth-"& money & " where username='" & membername & "'"
		conn.execute sqlbbs
		sql="update [KeHu] set ZiJin=ZiJin+" & money&",ZongZiJin=ZongZiJin+"&money&" where id="&MyUserID
		gp_conn.execute sql
		sucmess="<li>你已经把 "&money&" 元存进了你的股票帐户"
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
	KaiHu_Setting=gp_conn.execute("select top 1 KaiHu_Setting from GupiaoConfig order by id")(0)
	KaiHu_Setting=split(KaiHu_Setting,"|")
	KaiHu_Setting(0)=CLNG(KaiHu_Setting(0))
	KaiHu_Setting(1)=CLNG(KaiHu_Setting(1))
	KaiHu_Setting(2)=CLNG(KaiHu_Setting(2))
	KaiHu_Setting(3)=CLNG(KaiHu_Setting(3))
	KaiHu_Setting(4)=CLNG(KaiHu_Setting(4))
	KaiHu_Setting(6)=CLNG(KaiHu_Setting(6))
	if KaiHu_Setting(0)=0 then
		errmess="<li>股市暂停开户，请与管理员联系"
		call endinfo(1)
		exit sub
	end if

	sqlbbs="select Article,userCP,userEP,userWealth from [user] where username='"&membername&"'"
	set rsbbs=conn.execute(sqlbbs)	

	if rsbbs(0)<KaiHu_Setting(1) then
		errmess=errmess+"<li>本股市已限制了开户最少发帖数为"&KaiHu_Setting(1)&"，您现在不符合本要求，暂时不能开户"
		founderr=true
	end if
	if rsbbs(1)<KaiHu_Setting(2) then
		errmess=errmess+"<li>本股市已限制了开户最少魅力值为"&KaiHu_Setting(2)&"，您不符合本要求，暂时不能开户"
		founderr=true
	end if
	if rsbbs(2)<KaiHu_Setting(3) then
		errmess=errmess+"<li>本股市已限制了开户最少经验值为"&KaiHu_Setting(3)&"，您不符合本要求，暂时不能开户"
		founderr=true
	end if
	if rsbbs(3)<KaiHu_Setting(4) then
		errmess=errmess+"<li>本股市已限制了开户最少现金值为"&KaiHu_Setting(4)&"，您不符合本要求，暂时不能开户"
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
			sucmess="<li>股票帐户申请成功！<br><li><font color=#000099>由于你以前已经领取过股市资金，本次不再发放！</font> "
		else
			sql="insert into [KeHu](ZhangHao,ZiJin,ZongZiJin) values('" & membername & "',"&KaiHu_Setting(6)&","&KaiHu_Setting(6)&")"
			gp_conn.execute sql
			sucmess="<li>股票帐户申请成功！<br><li><font color=#000099>看到你是个菜鸟，股票市场便发慈悲送了<font color=#990000><b> "&KaiHu_Setting(6)&" </b></font> 元进了你的帐户里，乐吧？</font> "
		end if
		call endinfo(2)
	end if	
end sub
%>