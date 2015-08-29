<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="Z_bankconn.asp"-->
<!--#include file="Z_bankconfig.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
dim menu,savem,rsbank,rssql
dim lockac     '帐户是否被冻结  我来了 2002.10.23
dim smoney,saveday,time1,time2,daikuang,dkdu,daiday

menu=request.querystring("menu")
if menu="" then menu=0
stats="电子银行 银行营业大厅"

select case menu
	case 1
		stats="电子银行 社区二十大富翁"
		call nav()
		call head_var(0,0,"电子银行","Z_bank.asp")
	case 2
		stats="电子银行 银行二十大储户"
		call nav()
		call head_var(0,0,"电子银行","Z_bank.asp")
	case 8
		stats="电子银行 银行行长办公室"
		call nav()
		call head_var(0,0,"电子银行","Z_bank.asp")
	case 13
		stats="电子银行 浏览银行事件记录"
		call nav()
		call head_var(0,0,"电子银行","Z_bank.asp")		
	case else
		call nav()
		call head_var(2,0,"","")
end select

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入电子银行的权限，请先登录或者同管理员联系。"
	call dvbbs_error()
elseif cint(bank_setting(6))=1  and (not master) then '银行是否设置定时营业
	if DatePart("h", time)<cint(Chen_BusinessTimeSlice(0)) or DatePart("h", time)>=cint(Chen_BusinessTimeSlice(1)) then
		call BankClose()	
	else
		call start()
	end if
else
	call start()
end if
if founderr then  call dvbbs_error()
call activeonline()
call footer() 
'=========================================================
sub start()
	set rs=server.createobject("adodb.recordset")
    sql="select * from [bank] where username='"&membername&"' "
    rs.open sql,conn1,1,3
    if rs.eof then
		rs.addnew
		rs("username")=membername
		rs("bankuser_setting")="0,"&bank_setting(4)
		rs.update
		smoney=0
    end if
    daikuang=rs("daikuang")
	lockac=rs("lockac")
	smoney=rs("savemoney")
	bankuser_setting=split(rs("bankuser_setting"),",")
	
	'计算贷款信誉额
	if daikuang>0 then
    	dkdu=0
	elseif cint(bankuser_setting(0))=1 then
		dkdu=int(mymoney*bankuser_setting(1))         
	else
		dkdu=int(mymoney*bank_setting(4))	
	end if

    saveday=datediff("d",rs("date"),date())
    daiday=datediff("d",rs("dkdate"),date())
    rs.close
'-------------------------
%>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<tr>
		<th height=25><%=stats%></td>
	</tr>
	<tr>
		<td class=tablebody2 height=1>	
<%
	founderr=false
	
	call bankhead()
	
	select case menu
		case 0
			call main()					'银行营业大厅
		case 1
			call forum_top20()			'社区二十大富翁
		case 2
			call bank_top20()			'银行二十大储户   
		case 3
			call savemoney()			'存款	
		case 4
			call qumoney()				'取款
		case 5
			call zmoney()				'转帐
		case 6
			call daimoney()				'贷款
		case 7
			call hmoney()   			'还贷
		case 8
			call admin()				'银行管理页面
		case 9 
			call admin1()				'保存银行设置
		case 10
			call hmoney2()				'强制还款 
		case 11
			call jiangli()				'奖励
		case 12
			call savelogsetting()		'保存事件设置
		case 13
			call banklog()				'银行事件查看
		case else 
			call main() 											
	end select
%>
	</td></tr></table>
<%
	founderr=false
end sub   

'=======================================================================================
'-------------------------------------------主程序-----------------------------------
sub main()
     if daiday>daitian and daikuang>0 then
%>
		<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%"><tr><th height=26 >银行保安</th></tr><tr height=200><td align=center height=26 class=tablebody1><br><font color=red>由于你的贷款超过了尝还期限,现在强制执行！<%=hmoney2%></font><br><br></td></tr><tr><td align=center height=26 class="tablebody1"><a href="Z_bank.asp">返回银行大厅</a></td></tr></table>
<%   else  %>

<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
	<tr>
		<th height=26>欢迎来到电子银行管理您的财产</th>
	</tr>
	<tr>
		<td align=center height=26 class=tablebody2>银行目前的日利率是 <font color="#FF0000"><%=culi%>%</font>，贷款利率是 <font color="#FF0000"><%=daili%>%</font>，存款利率从第 <font color="#FF0000"><%=Chen_StartLixiDay%></font> 天起计算，当前银行储备资金为 <font color="#FF0000"><%=chubei%></font><br>贷款最长期限为 <font color="#FF0000"><%=daitian%></font> 天（到期将强制执行）；每次存款、取款自动结算利息</td>
	</tr>
<%if lockac then%>	
	<tr>
		<td align=center height=26 class=tablebody1><font color="#FF0000">您的银行账户已经被冻结,您不能进行取款、转账、贷款操作</font></td>
	</tr>   
<%
end if
if cint(bankstate)=1 then
%>    
	<form name="form1" method="post" action="Z_bank.asp?menu=3"><tr><td align=left height=30 class=tablebody1><%if mymoney>0 and cint(bank_setting(0))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%>　　　　<font face=Wingdings>v</font> 存款： <INPUT name=saving> <INPUT type=submit value=存入 name=submit> 你的现金为：<font color="#FF0000"><%=mymoney%></font> 元　存款利率为：<font color="#FF0000"><%=culi%>%</font> </td></tr></form>
	
	<form name="form2" method="post" action="Z_bank.asp?menu=4"><tr><td align=left height=30 class=tablebody1><%if lockac or smoney<=0 or cint(bank_setting(1))=0 then%><font color=red>×</font><%else%><font color=blue>√</font><%end if%>　　　　<font face=Wingdings>v</font> 取款： <INPUT name=draw> <INPUT type=submit value=取出 name=submit> 你现在的存款为：<font color="#FF0000"><%=smoney%></font> 元　你的利息为：<font color="#FF0000"><% if saveday<Chen_StartLixiDay then%>0<%else%><%=clng((formatnumber(smoney)*(formatnumber(culi)/100))*saveday)%><%end if%></font> 元 <%if smoney>0 then%>你的存款天数：<font color="#FF0000"><%=saveday%></font> 天 <%end if%></td></tr></form>
	
	<form name="form3" method="post" action="Z_bank.asp?menu=5"><tr><td align=left height=30 class=tablebody1><%if lockac or cint(bank_setting(3))=0 or (daikuang>0 and cint(bank_setting(5))=0)  then%><font color=red>×</font><%else%><font color=blue>√</font><%end if%>　　　　<font face=Wingdings>v</font> 转账： <INPUT name=trans> 元 给 <INPUT name=transTo> <INPUT type=submit value=转入 name=submit> 转账手续费：<font color="#FF0000"><%=zhuangli%>%</font> </td></tr></form>      
		  
	<form name="form4" method="post" action="Z_bank.asp?menu=6"><tr><td align=left height=30 class=tablebody1><%if lockac or daikuang>0 or cint(bank_setting(2))=0 then%><font color=red>×</font><%else%><font color=blue>√</font><%end if%>　　　　<font face=Wingdings>v</font> 贷款： <INPUT name=loan> <INPUT type=submit value=贷款 name=submit> 你的贷款额度为：<font color="#FF0000"><%=dkdu%></font> 元　贷款利率为：<font color="#FF0000"><%=daili%>%</font>
			  </td></tr></form>
	<%if daikuang>0 then%>
	<form name="form5" method="post" action="Z_bank.asp?menu=7"><tr><td align=left height=30 class=tablebody1>　　　　还清贷款： <INPUT type=submit value=还款 name=submit> 你现在的贷款额为：<font color="#FF0000"><%=daikuang%></font>　你贷款的利息为：<font color="#FF0000"><%if daikuang=0 then %>0<%else%><%=clng((daikuang*(formatnumber(culi)/100))*daiday)%></font> 元 你的贷款天数：<font color="#FF0000"><%=daiday%></font> 天<%end if%></td></tr></form>
	<%end if%>
<%else%>
	<tr><td align=center height=25 class=tablebody1><font face=Wingdings color=blue>v</font><font color=red> 银行结余/调整，暂时停止营业，如有问题请与管理员联系 </font><font face=Wingdings color=blue>v</font></td></tr>
<%end if%>
	<tr><td align=center height=25 class=tablebody2><font color=gray>银行营业时间：<% if cint(bank_setting(6))=0 then%>全天营业<%else%><%=Chen_BusinessTimeSlice(0)%>:00～<%=Chen_BusinessTimeSlice(1)%>:00<%end if%></font></td></tr>
</table>
<br>
<%
end if
end sub
'--------------------------------强制还款-------------------------------
function hmoney2()
    dim name
    if request.querystring("username")<>"" then
    	name=checkStr(trim(request.querystring("username")))
		if not master then 
			Errmsg=Errmsg+"<br>"+"<li>您没有执行强制还款操作的权限，请与管理员联系"
			call bank_err()
			exit function
		end if	
    else
    	name=membername
    end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from bank where username='"&name&"' "
	rs.open sql,conn1,1,3
	if rs.eof and rs.bof then 
		errmsg=errmsg+"<br>"+"<li>用户["&name&"]的银行账号不存在"
		call bank_err()
	else
		dim lixi,benjin,loan
		benjin=rs("daikuang")
		lixi=clng((daikuang*(formatnumber(daili)/100))*daiday)		
		loan=lixi+benjin
		rs("daikuang")=0
		rs("dkdate")=date()
		rs.update
		rs.close
		
		sql="select * from bankconfig"
		rs.open sql,conn1,1,3
		rs("chubei")=rs("chubei")+loan
		rs.update
		rs.close
		
		sql="select * from [user] where username='"&name&"' "
		rs.open sql,conn,1,3
		rs("userwealth")=rs("userwealth")-loan
		rs.update
		rs.close

		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content="<font color=red> <font color=blue>"&name&"</font> 被强制还清所有贷款，原贷款金额为："&benjin&" 元，还款金额为："&loan&"元</font>"
			call logs("管理","强制还款","银行行长")
		end if
		
		if request.querystring("username")<>"" then 
			sucmsg=sucmsg+"<br>"+"<li>用户["&name&"]的借款已经被强制还清，本金："&benjin&"元，利息："&lixi&"元，还款总金额为："&loan&"元"
			call bank_suc()
		else
			hmoney2="<br>"+"<font color=navy>您的借款已经被强制还清，本金："&benjin&"元，利息："&lixi&"元，还款总金额为："&loan&"元</font>"
		end if 
		
	end if 
end function 

'--------------------------------还贷程序-------------------------------
sub hmoney()
	if bankstate=0 then		'我来了 添加 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>银行暂停营业，请与管理员联系"
		call bank_err()
		exit sub	
	end if
		dim loan
		set rs=server.createobject("adodb.recordset")
		sql="select * from bank where username='"&membername&"' "
		rs.open sql,conn1,1,3
		set rsbank=server.createobject("adodb.recordset")
		rssql="select * from [user] where username='"&membername&"'"
		rsbank.open rssql,conn,1,3
		loan=clng((daikuang*(formatnumber(daili)/100))*daiday)+rs("daikuang")
        if rs("daikuang")=0 then
			rsbank.close
			Errmsg=Errmsg+"<br>"+"<li>您没有贷款啊"
	    	call bank_err()
		elseif loan>rsbank("userwealth") then
			rsbank.close
			rs("lockac")=true
			rs.update
			Errmsg=Errmsg+"<br>"+"<li>您没有那么多的钱还贷款"
			Errmsg=Errmsg+"<br>"+"<li>银行将会暂时冻结您的银行帐户，直到您还清您的贷款"
	    	call bank_err()		
			content="不够钱还款，银行账号自动被冻结"
			call logs("银行","偿还贷款",membername)	
        else
			rsbank.close
			rs("daikuang")=0
			rs("dkdate")=date()
			'rs("lockac")=false
			rs.update
			set rs=server.createobject("adodb.recordset")
			sql="select * from bankconfig"
			rs.open sql,conn1,1,3
			rs("chubei")=rs("chubei")+loan
			rs.update
			rs.close
			set rs=server.createobject("adodb.recordset")
			sql="select * from [user] where username='"&membername&"' "
			rs.open sql,conn,1,3
			rs("userwealth")=rs("userwealth")-loan
			rs.update
			rs.close
			
			if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
				if cint(log_setting(4))=1 or (cint(log_setting(4))=2 and loan>=clng(log_setting(5))) then
					content="还清所有贷款，原贷款金额为："&daikuang&" 元，还款金额为："&loan&"元"
					call logs("银行","偿还贷款",membername)
					sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
				end if	
			end if
					   
			sucmsg=sucmsg+"<br>"+"<li>您的借款已经全部还清，原贷款金额为："&daikuang&" 元，还款金额为："&loan&"元"
			call bank_suc() 		   
		end if
end sub

'--------------------------------贷款程序--------------------------------
sub daimoney()
	if bankstate=0 then			'我来了 添加 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>银行暂停营业，请与管理员联系"
		call bank_err()
		exit sub	
	end if
	if cint(bank_setting(2))=0 then			'我来了 添加 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>银行贷款服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub	
	end if	
	if lockac then 			'我来了 添加
		Errmsg=Errmsg+"<br>"+"<li>对不起，您不能进行该操作"
		Errmsg=Errmsg+"<br>"+"<li>您的银行账户已经被冻结，请与管理员联系"
    	call bank_err()
		exit sub	
	end if
	
	dim loan
    if request.form("loan")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入贷款金额"
    	founderr=true
	elseif not isnumeric(request.form("loan")) then
		Errmsg=Errmsg+"<br>"+"<li>贷款金额必须是数字"
    	founderr=true		
	elseif request.form("loan")<=0 then
		Errmsg=Errmsg+"<br>"+"<li>贷款金额必须是正数"
    	founderr=true
	else
		loan=int(request.form("loan"))
		if loan>dkdu then
			Errmsg=Errmsg+"<br>"+"<li>您贷款的金额超出了您的贷款信用额度"
			founderr=true
		elseif loan>chubei then
			Errmsg=Errmsg+"<br>"+"<li>您的贷款金额太大了，银行储备不足，请与管理员联系"
			founderr=true			
		end if 		
	end if		
			
	if founderr then
		call bank_err()		
    else
		set rs=server.createobject("adodb.recordset")
		sql="select * from bank where username='"&membername&"' "
		rs.open sql,conn1,1,3
        if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>您的银行账号不存在，请从有效连接进入"
			call bank_err()
 		else 
			rs("daikuang")=rs("daikuang")+loan
			rs("dkdate")=date()
			rs.update
			rs.close
			set rs=server.createobject("adodb.recordset")
			sql="select * from bankconfig"
			rs.open sql,conn1,1,3
			rs("chubei")=rs("chubei")-loan
			rs.update
			rs.close
			set rs=server.createobject("adodb.recordset")
			sql="select * from [user] where username='"&membername&"' "
			rs.open sql,conn,1,3
			rs("userwealth")=rs("userwealth")+loan
			rs.update
			rs.close
			
			if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
				if cint(log_setting(4))=1 or ( cint(log_setting(4))=2 and loan>=clng(log_setting(5)) ) then
					content="向银行贷款，贷款金额为："&loan&" 元，贷款期限为："&daitian&"天"
					call logs("银行","向银行贷款",membername)
					sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
				end if	
			end if
								
			sucmsg=sucmsg+"<br>"+"<li>您的贷款手续已经办理好了，贷款金额为："&loan&"元，贷款期限为："&daitian&"天"
			call bank_suc()               
		end if
	end if
end sub

'--------------------------------转帐程序------------------------------
sub zmoney()
	if bankstate=0 then			'我来了 添加 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>银行暂停营业，请与管理员联系"
		call bank_err()
		exit sub	
	end if
	if cint(bank_setting(3))=0 then			'我来了 添加 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>银行转帐服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub	
	end if
	if lockac then           '我来了 添加
		Errmsg=Errmsg+"<br>"+"<li>对不起，您不能进行该操作"
		Errmsg=Errmsg+"<br>"+"<li>您的银行账户已经被冻结，请与管理员联系"
    	call bank_err()
		exit sub	
	end if
	if daikuang>0 and cint(bank_setting(5))=0 then		'我来了 添加 2002.12.03
		Errmsg=Errmsg+"<br>"+"<li>银行规定贷款之后不能转帐，请与管理员联系"
    	call bank_err()
		exit sub	
	end if
	
    if request.form("trans")="" or (not isnumeric(request.form("trans"))) then
		Errmsg=Errmsg+"<br>"+"<li>请正确输入转帐金额"
		founderr=true
	elseif request.form("trans")<=0 then
		Errmsg=Errmsg+"<br>"+"<li>转帐金额必须是正数"
		founderr=true
	end if		
    if trim(request.form("transto"))="" then
		Errmsg=Errmsg+"<br>"+"<li>您要转帐到哪个账号上？"
		founderr=true
	end if
	if founderr then
		call bank_err()
    else
	   	dim trans,transto
	   	trans=clng(request.form("trans"))
	   	transto=checkStr(trim(request.form("transto")))
	   
	   	set rs=server.createobject("adodb.recordset")
	   	set rs=conn.execute("select username from [user] where username='"&transto&"'")
	   	if rs.bof and rs.eof then
			Errmsg=Errmsg+"<br>"+"<li>论坛上没有["&transto&"]这个用户!"
    		call bank_err()
			rs.close
			exit sub
		else
	   		transto=rs(0)
			rs.close
		end if
        if myarticle<100 then 
  Errmsg=Errmsg+""+"<li>银行转帐服务规定发帖数超过100才可以转帐"
  call bank_err()
  exit sub 
 end if
	   	sql="select * from bank where username='"&membername&"' "
	   	rs.open sql,conn1,1,3
        if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>您的银行账号不存在，请从有效连接进入"
			call bank_err()
 		else 
			if rs("savemoney") < trans then
				Errmsg=Errmsg+"<br>"+"<li>您在银行中没有那么多的存款"
				call bank_err()
			elseif rs("savemoney")<int(trans*(1+formatnumber(zhuangli)/100)+0.999999) then
				Errmsg=Errmsg+"<br>"+"<li>您在银行存款余额不足于支付转帐手续费"
				call bank_err()		
			else
				rs("savemoney")=rs("savemoney")-int(trans*(1+formatnumber(zhuangli)/100)+0.999999)
				rs.update
				rs.close
				set rs=server.createobject("adodb.recordset")
				sql="select * from bank where username='"&transto&"' "
				rs.open sql,conn1,1,3
				if rs.eof then
					rs.addnew
					rs("username")=transto
					rs("bankuser_setting")="0,"&bank_setting(4)
				end if    
				rs("savemoney")=rs("savemoney")+trans
				rs.update
				rs.close
				set rs=server.createobject("adodb.recordset")
				sql="select * from bankconfig"
				rs.open sql,conn1,1,3
				rs("chubei")=rs("chubei")+int(trans*(formatnumber(zhuangli)/100)+0.999999)
				rs.update
				rs.close
				
				if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
					if cint(log_setting(4))=1 or (cint(log_setting(4))=2 and trans>=clng(log_setting(5))) then
						content="转入账号：<font color=blue>"&HTMLEncode(transto)&"</font>, 转帐金额为："&trans&"元, 转帐手续费："&int(trans*(formatnumber(zhuangli)/100)+0.999999)&"元"
						call logs("银行","办理账号手续",membername)
						sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
					end if	
				end if
							
				sucmsg=sucmsg+"<br>"+"<li>您的转帐手续已经办理好了"
				sucmsg=sucmsg+"<br>"+"<li>转入账号：<font color=blue>"&HTMLEncode(transto)&"</font>, 转帐金额为："&trans&"元, 转帐手续费："&int(trans*(formatnumber(zhuangli)/100)+0.999999)&"元"
					dim sender,title,body,sql2
                 sender=membername
                 title="转帐通知!"
                  body=membername&"通过社区电子银行的转帐系统送 "&trans&" 元给你，请注意查收!"
                  sql2="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&HTMLEncode(transto)&"','"&sender&"','"&title&"','"&body&"',Now(),0,1)"
                  conn.Execute(sql2)
			call bank_suc()				   
			end if
		end if
	end if
end sub

'--------------------------------取款程序--------------------------------
sub qumoney()
	if bankstate=0 then			'我来了 添加 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>银行暂停营业，请与管理员联系"
		call bank_err()
		exit sub	
	end if
	if cint(bank_setting(1))=0 then			'我来了 添加 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>银行提款服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub	
	end if
	if lockac then 			'我来了 添加
		Errmsg=Errmsg+"<br>"+"<li>对不起，您不能进行该操作"
		Errmsg=Errmsg+"<br>"+"<li>您的银行账户已经被冻结，请与管理员联系"
    	call bank_err()
		exit sub	
	end if
    if request.form("draw")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入取款金额"
		founderr=true
	elseif not isnumeric(request.form("draw")) then
		Errmsg=Errmsg+"<br>"+"<li>取款金额必须是数字"
		founderr=true
	elseif request.form("draw")<=0 then
		Errmsg=Errmsg+"<br>"+"<li>取款金额必须是正数"
		founderr=true			
    else
		dim drawm
		drawm=int(request.form("draw"))	
		if drawm>chubei then
			Errmsg=Errmsg+"<br>"+"<li>您的取款金额太大了，银行储备不足，请与管理员联系"
			founderr=true			
		end if 
	end if	
	if founderr then
		call bank_err()	
	else
		dim saveli
		set rs=server.createobject("adodb.recordset")
		sql="select * from bank where username='"&membername&"' "
		rs.open sql,conn1,1,3
        if rs.eof and rs.bof then
			Errmsg=Errmsg+"<br>"+"<li>您的银行账号不存在，请从有效连接进入"
			call bank_err()		
        elseif drawm>rs("savemoney") then
			Errmsg=Errmsg+"<br>"+"<li>您在银行中没有那么多的存款"
	    	call bank_err()	
        else
			if saveday<Chen_StartLixiDay then
				saveli=0
			else
				saveli=clng((smoney*(formatnumber(culi)/100))*saveday)
			end if
			rs("savemoney")=rs("savemoney")-drawm+saveli
			rs("date")=date()
			rs.update
			rs.close

			sql="select userwealth from [user] where username='"&membername&"' "
			rs.open sql,conn,1,3
			if not(rs.eof and rs.bof) then
				rs(0)=rs(0)+drawm
				rs.update
			end if
			rs.close
			
			sql="select chubei from bankconfig"
			rs.open sql,conn1,1,3
			rs(0)=rs(0)-drawm
			rs.update
			rs.close 
				   
			if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
				if cint(log_setting(4))=1 or (cint(log_setting(4))=2 and drawm>=clng(log_setting(5))) then
					content="向银行提取"&drawm&"元"
					call logs("银行","提款",membername)
					sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
				end if	
			end if
			
		   sucmsg=sucmsg+"<br>"+"<li>您的存款存款利息已经存入您的账号，您得到的利息是："&saveli&"元" 					   
		   sucmsg=sucmsg+"<br>"+"<li>您的取款手续已经办理好了，取款金额为："&drawm&"元"
		   call bank_suc()				   
 end if
end if
end sub

'--------------------------------------存款程序------------------------------------
sub savemoney()
	if bankstate=0 then			'我来了 添加 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>银行暂停营业，请与管理员联系"
		call bank_err()
		exit sub	
	end if
	if cint(bank_setting(0))=0 then			'我来了 添加 2002.11.30
		Errmsg=Errmsg+"<br>"+"<li>银行存款服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub	
	end if
    if request.form("saving")="" or (not isnumeric(request.form("saving"))) then
		Errmsg=Errmsg+"<br>"+"<li>请正确输入存款金额"
    	call bank_err()
    else
		savem=clng(request.form("saving"))
		set rs=server.createobject("adodb.recordset")
		sql="select * from [user] where username='"&membername&"' "
		rs.open sql,conn,1,3
       if savem>rs("userwealth") then
			Errmsg=Errmsg+"<br>"+"<li>您没有那么多的钱"
			call bank_err()	   		
 	   elseif savem<0 then 
			Errmsg=Errmsg+"<br>"+"<li>请正确输入存款金额"
    		call bank_err()
   	   else 
           set rsbank=server.createobject("adodb.recordset")
           rssql = "select * from bank where username='"&membername&"' "
           rsbank.open rssql,conn1,1,3
           if rsbank.eof then
              rsbank.addnew
              rsbank("username")=membername
              rsbank("savemoney")=savem
              rsbank.update
              rsbank.close
                set rsbank=server.createobject("adodb.recordset")
                rssql = "select * from bankconfig"
                rsbank.open rssql,conn1,1,3
                rsbank("chubei")=rsbank("chubei")+savem
                rsbank.update
                rsbank.close
              rs("userwealth")=rs("userwealth")-savem
              rs.update
              rs.close
           else
				dim saveli
				if saveday<Chen_StartLixiDay then
					saveli=0
				else
					saveli=clng((smoney*(formatnumber(culi)/100))*saveday)
				end if
				rsbank("savemoney")=rsbank("savemoney")+savem+saveli
				rsbank("date")=date()
				rsbank.update
				rsbank.close
                set rsbank=server.createobject("adodb.recordset")
                rssql = "select * from bankconfig"
                rsbank.open rssql,conn1,1,3            
                rsbank("chubei")=rsbank("chubei")+savem
                rsbank.update
                rsbank.close
				rs("userwealth")=rs("userwealth")-savem
				rs.update
				rs.close
           end if
		   
			if cint(log_setting(0))=1 and cint(log_setting(4))<>0 then
				if cint(log_setting(4))=1 or (cint(log_setting(4))=2 and savem>=clng(log_setting(5))) then
					content="向银行存入"&savem&"元"
					call logs("银行","存款",membername)
					sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
				end if	
			end if
			sucmsg=sucmsg+"<br>"+"<li>您的存款存款利息已经存入您的账号，您得到的利息是："&saveli&"元" 		   
		   	sucmsg=sucmsg+"<br>"+"<li>您的钱成功的存入银行，本次存入："&savem&"元"
		   	call bank_suc()
		end if
	end if
end sub

'-------------------------------------------论坛TOP20大富翁排行-------------------------------------
sub forum_top20()
    set rs=server.createobject("adodb.recordset")
    sql="select * from [user] order by userWealth desc"
    rs.open sql,conn,1,3
%>

<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%">
<tr> 
	<th>用户名</th> 
	<th>Email</th>
	<th>OICQ</th>
	<th>主页</th>
	<th>短消息</th>
	<th>注册时间</th>
	<th>等级状态</th>
	<th>发帖总数</th>
	<th>金额</th>
</tr>	
<%
dim i
       for i=1 to 20
%>
<tr>
<td align=center class=tablebody2><a href="dispuser.asp?name=<%=htmlencode(rs("username"))%>" target=_blank><%=rs("username")%></a></td>
<td align=center class=tablebody1><a href=mailto:<%=rs("useremail")%>><img border=0 src="pic/email.gif"></a></td>
<td align=center class=tablebody2> 
<%if rs("oicq")="" or isnull(rs("oicq")) then%>
没有       
<%else%>      
<a href=http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<%=rs("oicq")%> target=_blank><img src="pic/oicq.gif" alt="查看 OICQ:<%=rs("oicq")%> 的资料" border=0 width=16 height=16></a>       
<%end if%>      
</td>      
<td align=center class=tablebody1>       
<%if rs("homepage")="" or isnull(rs("homepage")) then%>      
没有       
<%else%>      
<a href=<%=rs("homepage")%> target=_blank><img border=0 src=pic/homepage.gif></a>       
<%end if%>      
</td>      
<td align=center class=tablebody2><a href=usersms.asp?action=new&touser=<%=htmlencode(rs("username"))%> target=_blank><img src=pic/message.gif border=0></a></td>      
<td align=center class=tablebody1><%=rs("addDate")%></td>      
<td align=center class=tablebody2><%=rs("userclass")%><br>      
</td>      
<td align=center class=tablebody1><%=rs("article")%></td>      
<td align=center class=tablebody2><%=rs("userwealth")%></td>      
</tr>      
<%      
           rs.movenext      
         if rs.eof then      
             exit for      
         end if      
       next      
%>      
</table>

</td></tr>  
<tr><td align=center colspan=9 height=26 class=tablebody1><a href=Z_bank.asp>返回银行首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>返回上一页</a></td></tr>    
<%      
end sub       
     
'---------------------------------二十大储户排行-----------------------------------      
sub bank_top20()      
    set rs=server.createobject("adodb.recordset")      
    sql="select * from bank order by savemoney desc"      
    rs.open sql,conn1,1,3      
%>      

<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%">    
<tr>       
<th>帐户</th>      
<th>存款</th>  
<th>存款利息</th>
<th>存款日期</th>      
<th>贷款</th>  
<th>贷款利息</th>    
<th>贷款日期</th>
<th width="65">冻结帐户</th>  
</tr>      
<%      
dim i,savedaytmp,daidaytmp    
       for i=1 to 20  
			savedaytmp=datediff("d",rs("date"),date())
			daidaytmp=datediff("d",rs("dkdate"),date())	       
%>      
			<tr>       
			<td class=tablebody2 align=center><a href="dispuser.asp?name=<%=rs("username")%>" target=_blank><%=rs("username")%></a></td>      
			<td class=tablebody1 align=center><%=rs("savemoney")%></td>     
			<td class=tablebody1 align=center><%if savedaytmp<Chen_StartLixiDay then %>0<%else%><font color="#0066FF"><%=clng((formatnumber(rs("savemoney"))*(formatnumber(culi)/100))*savedaytmp)%></font><%end if%></td> 
			<td class=tablebody2 align=center><%=rs("date")%></td> 
			<td class=tablebody1 align=center><%=rs("daikuang")%></td> 
			<td class=tablebody1 align=center><%if rs("daikuang")=0 then %>0<%else%><font color="#0066FF"><%=clng((formatnumber(rs("daikuang"))*(formatnumber(daili)/100))*daidaytmp)%></font><%end if%></td> 
			<td class=tablebody2 align=center><%=rs("dkdate")%></td>
			<td class=tablebody1 align=center><%if rs("lockac") then%><font color="#FF0000">是</font><%else%>否<%end if%></td>     
			</tr>      
<%      
           rs.movenext      
         if rs.eof then      
             exit for      
         end if      
       next      
%>      
</table>
<br>
</td></tr>  
<tr><td align=center colspan=8 height=26 class=tablebody1><a href=Z_bank.asp>返回银行首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>返回上一页</a></td></tr>      
<%      
end sub 

'----------------------------银行基本设置--------------------------
sub admin()       
      
if not master then   
	Errmsg=Errmsg+"<br>"+"<li>您没有进入银行基本设置的权限。"   
  	call bank_err()      
else      
%>      
<br>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%">      
<tr><th colspan=4 height=20 align=left>银行管理中心---银行基本设置</th></tr>  
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr>    
<form name="form2" method="post" action="Z_bank.asp?menu=9">      
<tr>       
	<td align=right width="20%" class=tablebody2>存款的利息：</td>      
	<td align=left width="30%" class=tablebody1><INPUT name=saveli value="<%=culi%>"> %</td>   
	<td align=right width="20%" class=tablebody2>存款服务：</td> 
	<td align=center width="30%" class=tablebody1>
	<input type=radio name="canCunkuan" value=1 <%if cint(bank_setting(0))=1 then%>checked<%end if%>>开通&nbsp;
	<input type=radio name="canCunkuan" value=0 <%if cint(bank_setting(0))=0 then%>checked<%end if%>>关闭&nbsp;
	</td> 
</tr>      
<tr>       
	<td align=right class=tablebody2>贷款的利息：</td>      
	<td align=left class=tablebody1><INPUT name=daili value="<%=daili%>"> %</td>  
	<td align=right class=tablebody2>贷款服务：</td> 
	<td align=center class=tablebody1>
	<input type=radio name="canDaikuan" value=1 <%if cint(bank_setting(2))=1 then%>checked<%end if%>>开通&nbsp;
	<input type=radio name="canDaikuan" value=0 <%if cint(bank_setting(2))=0 then%>checked<%end if%>>关闭&nbsp;
	</td>    
</tr>      
<tr>       
	<td align=right class=tablebody2>银行的储备资金：</td>      
	<td align=left class=tablebody1><INPUT name=chubei value="<%=chubei%>"> 元</td>  
	<td align=right class=tablebody2>取款服务：</td> 
	<td align=center class=tablebody1>
	<input type=radio name="canQukuan" value=1 <%if cint(bank_setting(1))=1 then%>checked<%end if%>>开通&nbsp;
	<input type=radio name="canQukuan" value=0 <%if cint(bank_setting(1))=0 then%>checked<%end if%>>关闭&nbsp;
	</td>      
</tr>      
<tr>       
	<td align=right class=tablebody2>转帐的手续费：</td>      
	<td align=left class=tablebody1><INPUT name=zhuangli value="<%=zhuangli%>"> %</td> 
	<td align=right class=tablebody2>转帐服务：</td> 
	<td align=center class=tablebody1>
	<input type=radio name="canZhuanzang" value=1 <%if cint(bank_setting(3))=1 then%>checked<%end if%>>开通&nbsp;
	<input type=radio name="canZhuanzang" value=0 <%if cint(bank_setting(3))=0 then%>checked<%end if%>>关闭&nbsp;
	</td>	     
</tr>      
<tr>       
	<td align=right class=tablebody2>最长贷款天数：</td>      
	<td align=left class=tablebody1><INPUT name=daitian value="<%=daitian%>"> 天</td>    
	<td align=right class=tablebody2><font color=red>银行状态：</font></td> 
	<td align=center class=tablebody1>
	<input type=radio name="bankstate" value=1 <%if cint(bankstate)=1 then%>checked<%end if%>><font color=red>营业</font>&nbsp;
	<input type=radio name="bankstate" value=0 <%if cint(bankstate)=0 then%>checked<%end if%>><font color=red>关闭</font>&nbsp; 
	</td> 	  
</tr> 
<tr>       
	<td align=right class=tablebody2>计息起始天数：</td>      
	<td align=left class=tablebody1><INPUT name="StartLixiDay" value="<%=Chen_StartLixiDay%>"> 天</td>    
	<td align=right class=tablebody2>营业时间</td> 
	<td align=center class=tablebody1>
	<input type=radio name="BusinessHours" value=0 <%if cint(bank_setting(6))=0 then%>checked<%end if%>>全天候&nbsp;
	<input type=radio name="BusinessHours" value=1 <%if cint(bank_setting(6))=1 then%>checked<%end if%>>定时&nbsp; 	
	<INPUT name="BusinessTimeSlice" value="<%=bank_setting(7)%>" size=10>
	</td> 	
</tr>   
<tr>       
	<td align=right class=tablebody2>默认信誉系数：</td>      
	<td align=left class=tablebody1><INPUT name=creditvalue value="<%=bank_setting(4)%>"></td>    
	<td align=right class=tablebody2>有贷款是否允许转帐</td> 
	<td align=center class=tablebody1>
	<input type=radio name="candaikuantrans" value=1 <%if cint(bank_setting(5))=1 then%>checked<%end if%>>是&nbsp;
	<input type=radio name="candaikuantrans" value=0 <%if cint(bank_setting(5))=0 then%>checked<%end if%>>否&nbsp; 	
	</td> 	  
</tr>     
<tr>      
	<td align=center colspan=4 class=tablebody1><INPUT type=submit value=修改 name=submit></td>      
</tr>      
</form> 
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr> 
<form name="form0" method="post" action="Z_bank.asp?menu=12">	
<tr><th colspan=4 height=20 align=left>银行管理中心---银行日志设置 (记录银行的各项操作)</th></tr>
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr> 
<tr>       
	<td class=tablebody2><font color=red>银行操作事件：</font></td>      
	<td class=tablebody1>
	<input type=radio name="logstate" value=1 <%if cint(log_setting(0))=1 then%>checked<%end if%>><font color=red>记录</font>&nbsp;
	<input type=radio name="logstate" value=0 <%if cint(log_setting(0))=0 then%>checked<%end if%>><font color=red>不记录</font>&nbsp; 	
	</td>    
	<td class=tablebody2>记录服务事件(特殊服务)：</td> 
	<td class=tablebody1>
	<input type=radio name="logfuwu" value=1 <%if cint(log_setting(1))=1 then%>checked<%end if%>>是&nbsp; 
	<input type=radio name="logfuwu" value=0 <%if cint(log_setting(1))=0 then%>checked<%end if%>>否&nbsp; 
	</td> 	  
</tr> 
<tr>       
	<td class=tablebody2>保存最新事件条数：<br>(保留所有事件填<font color=red>0</font>)</td>      
	<td class=tablebody1>
	&nbsp;<input type=text name="logrecordcount" value=<%=log_setting(2)%>>
	</td>    
	<td class=tablebody2>记录所有管理操作事件：</td> 
	<td class=tablebody1>
	<input type=radio name="logadmin" value=1 <%if cint(log_setting(3))=1 then%>checked<%end if%>>是&nbsp; 
	<input type=radio name="logadmin" value=0 <%if cint(log_setting(3))=0 then%>checked<%end if%>>否&nbsp; 
	</td> 	  
</tr>
<tr>       
	<td class=tablebody2>记录储户事件(银行事件)：</td> 
	<td class=tablebody1 colspan=3>    
	<input type=radio name="loguse" value=1 <%if cint(log_setting(4))=1 then%>checked<%end if%>>全部&nbsp; 
	<input type=radio name="loguse" value=2 <%if cint(log_setting(4))=2 then%>checked<%end if%>>金额大于
	<input type=text name="logminmoney" size=10 value=<%=log_setting(5)%>> 元的操作
	<input type=radio name="loguse" value=0 <%if cint(log_setting(4))=0 then%>checked<%end if%>>不记录 &nbsp;
	</td> 	  
</tr> 
<tr>   
	<td class=tablebody2>用户可否查看事件记录：</td> 
	<td class=tablebody1 colspan=3>    
	<input type=radio name="loguselook" value=1 <%if cint(log_setting(6))=1 then%>checked<%end if%>>可以&nbsp; 
	<input type=radio name="loguselook" value=0 <%if cint(log_setting(6))=0 then%>checked<%end if%>>不可以&nbsp; 
	</td> 	  
</tr> 
<tr>      
	<td align=center colspan=4 class=tablebody1><INPUT type=submit value=修改 name=submit></td>      
</tr> 
</form>   	
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr> 		 
<tr><th colspan=4 height=20 align=left>银行管理中心---银行贷款管理 (只有超过还款日期的帐户才列出)</th></tr>    
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr> 
<%      
    set rs=server.createobject("adodb.recordset")      
    sql="select * from bank where daikuang>0 and datediff('d',dkdate,date())>"&daitian
    rs.open sql,conn1,1,3      
%>      
<tr>      
	<th>帐户名</th>      
	<th>贷款金额</th>   
	<th>贷款天数</th>    
	<th>操   作</th>      
</tr>      
<%
if  rs.eof and rs.bof then
%>
	<tr>      
		<td align=center colspan=4 class=tablebody2>暂时还没有超过还款日期的帐户</td>      
	</tr> 
<%	
else     
	do while not rs.eof      
%>      
	<tr>      
	<td align=center class=tablebody2><a href="dispuser.asp?name=<%=rs("username")%>" target=_blank><%=rs("username")%></a></td>      
	<td align=center class=tablebody1><%=rs("daikuang")%></td>     
	<td align=center class=tablebody1><%=datediff("d",rs("dkdate"),date())%></td>  
	<td align=center class=tablebody2><a href="Z_bank.asp?menu=10&username=<%=rs("username")%>">强制还款</a></td>      
	</tr>      
<%       
	rs.movenext      
	loop 
end if	
%> 
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr>      
<tr height=22><th colspan=4 height=20 align=left>银行管理中心---银行奖励管理</th></tr> 
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr>
<form name="form1" method="post" action="Z_bank.asp?menu=11">      
<tr>      
<td align=right class=tablebody2>奖励对象：</td>      
<td align=left colspan=3 class=tablebody1><select name="name">      
    <option value="20" selected>所有管理员</option>
    <option value="19">所有贵宾</option>
    <option value="18">所有超级版主</option>
     <option value="17">所有版主</option>
    <option value="1">所有会员</option>
     <option value="24">所有VIP会员</option>
    <option value="23">登录20次以上的会员</option>
  </select>  奖励单个会员请填写：<INPUT name=username></td></tr>      
<tr>      
<td align=right class=tablebody2>奖励金额：</td>      
<td align=left colspan=3 class=tablebody1><INPUT name=money></td>      
</tr>      
<tr>      
<td align=right class=tablebody2>奖励理由：</td>      
<td align=left colspan=3 class=tablebody1><INPUT name=liyou size="50"></td>      
</tr>      
<tr>      
<td align=center colspan=4 class=tablebody1><INPUT type=submit value=奖励 name=submit></td>      
</tr>      
</form> 
	<tr>      
		<td align=center colspan=4 class=tablebody1 height=10></td>      
	</tr>      
<tr  height=22><td align=center colspan=4  height=26 class=tablebody2><a href=Z_bank.asp>返回银行首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>返回上一页</a></td></tr>      
</table>  
<br>    
<%       
end if      
end sub 

sub admin1() 
	dim saveli 
	if not master then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您没有进入银行基本设置的权限。"      
	else
		if request.form("saveli")="" or (not isnumeric(request.form("saveli"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>存款利息必须为数字"
		elseif request.form("saveli")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>存款利息不能为负数"
		end if	 
		if request.form("daili")="" or (not isnumeric(request.form("daili"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>贷款利息必须为数字" 
		elseif request.form("daili")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>贷款利息不能为负数"		
		end if	
		if request.form("chubei")="" or (not isnumeric(request.form("chubei"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>银行储备资金必须为数字" 
		elseif request.form("chubei")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>银行储备资金不能为负数"			
		end if	
		if request.form("zhuangli")="" or (not isnumeric(request.form("zhuangli"))) then			
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>转帐手续费必须为数字" 	
		elseif request.form("zhuangli")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>转帐手续费不能为负数"		
		end if
		if request.form("daitian")="" or (not isnumeric(request.form("daitian"))) then			
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>最长贷款天数必须为数字" 	
		elseif request.form("daitian")<=0 or int(request.form("daitian"))-request.form("daitian")<>0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>最长贷款天数不能必须为正整数"		
		end if
		if request.form("creditvalue")="" or (not isnumeric(request.form("creditvalue"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>信誉系数必须为数字"
		elseif request.form("creditvalue")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>信誉系数不能为负数"		
		end if
		if request.form("StartLixiDay")="" or (not isnumeric(request.form("StartLixiDay"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>请输入计息起始天数，或者您输入的不是数字"
		elseif request.form("StartLixiDay")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>计息起始天数不能为负数"		
		end if
		if request.form("BusinessTimeSlice")="" then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>请输入营业起始时间"
		else
			Chen_BusinessTimeSlice=split(request.form("BusinessTimeSlice"),"||")
			if ubound(Chen_BusinessTimeSlice)<>1 then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>营业起始时间输入不正确，格式为：开始小时数||结束小时数"
			elseif not (isnumeric(Chen_BusinessTimeSlice(0)) and isnumeric(Chen_BusinessTimeSlice(1))) then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>营业起始时间的'||'符号两边必须是数字"
			elseif cint(Chen_BusinessTimeSlice(0))<0 or  cint(Chen_BusinessTimeSlice(0))>24 then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>开始营业时间范围不正确，请输入0～24之间的整数"
			elseif cint(Chen_BusinessTimeSlice(1))<0 or  cint(Chen_BusinessTimeSlice(1))>24 then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>结束营业时间范围不正确，请输入0～24之间的整数"
			elseif cint(Chen_BusinessTimeSlice(1))<cint(Chen_BusinessTimeSlice(0)) then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>开始营业时间不能大于结束营业时间"												
			end if			
		end if		
	end if	
	if founderr then      
		call bank_err()
	else
		saveli=formatnumber(request.form("saveli"))      
		daili=formatnumber(request.form("daili"))      
		chubei=int(request.form("chubei"))      
		zhuangli=formatnumber(request.form("zhuangli"))      
		daitian=int(request.form("daitian"))      
		set rs=server.createobject("adodb.recordset")      
		sql="select * from [bankconfig]"      
		rs.open sql,conn1,1,3      
		rs("savedayli")=saveli      
		rs("ddayli")=daili      
		rs("chubei")=chubei      
		rs("zzli")=zhuangli      
		rs("dkday")=daitian
		rs("state")=request.form("bankstate")        '我来了 添加 2002.11.30
		rs("bank_setting")=request.Form("canCunkuan") & "," & request.Form("canQukuan")& "," & request.Form("canDaikuan")& "," & request.Form("canZhuanzang")& "," & request.form("creditvalue")& "," & request.form("candaikuantrans")& "," & request.form("BusinessHours")& "," & request.form("BusinessTimeSlice")
		  
		rs("StartLixiDay")=request.form("StartLixiDay")		'绿水青山 2003.1.6
		rs.update      
		rs.close 
		    
		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content="调整银行基本设置"
			call logs("管理","银行管理中心",membername)
			sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
		end if
				
	   sucmsg=sucmsg+"<br>"+"<li>银行基本信息修改完成"
	   call bank_suc()		  
	end if      
end sub  
     
'-----------------------------奖励-------------------------------       
sub jiangli()       
	dim money,liyou,classid,rs2,sql2  
	if not master then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您没有执行银行奖励的权限。"      
	else	
		if request.form("money")="" or (not isnumeric(request.form("money"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>奖励金额必须为数字"	  
		elseif request.form("money")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>奖励金额不能为负数"		
		end if
		if trim(request.form("liyou"))="" then
				founderr=true
				Errmsg=Errmsg+"<br>"+"<li>请输入奖励理由"	
		end if 
	end if	
	
	if founderr then
		call bank_err()
		exit sub
	end if	  
		dim classuser  '奖励单个用户
		money=int(request.form("money"))      
		liyou=trim(request.form("liyou"))
		if request.form("username") = "" then      
			classid=int(request.form("name"))     
			if classid=20 then      
				set rs=server.createobject("adodb.recordset")      
				sql="select userwealth,username from [user] where usergroupid=1"      
				rs.open sql,conn,1,3      
				do while not rs.eof      
					rs(0)=rs(0)+money      
					rs.update      
						set rs2=server.createobject("adodb.recordset")      
						sql2="select * from [message]"      
						rs2.open sql2,conn,1,3      
						rs2.addnew      
						rs2("sender")="银行行长"      
						rs2("incept")=rs(1)      
						rs2("title")="社区银行给您奖励"+cstr(money)+"元"      
						rs2("content")="社区很行给于您"+cstr(money)+"元奖励。奖励理由："+liyou       
						rs2("flag")=0      
						rs2("issend")=1      
						rs2.update      
						rs2.close      
                rs.movenext         
            loop      
         elseif classid=19 then      
                     set rs=server.createobject("adodb.recordset")      
            sql="select userwealth,username from [user] where usergroupid=8"      
            rs.open sql,conn,1,3      
            do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="银行行长"      
					rs2("incept")=rs(1)      
					rs2("title")="社区银行给您奖励"+cstr(money)+"元"      
					rs2("content")="社区很行给于您"+cstr(money)+"元奖励。奖励理由："+liyou       
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop  
            elseif classid=18 then      
                set rs=server.createobject("adodb.recordset")      
				sql="select userwealth,username from [user] where usergroupid=2"      
				rs.open sql,conn,1,3      
				do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="银行行长"      
					rs2("incept")=rs(1)      
					rs2("title")="社区银行给您奖励"+cstr(money)+"元"      
					rs2("content")="社区很行给于您"+cstr(money)+"元奖励。奖励理由："+liyou       
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop      
    elseif classid=17 then      
            set rs=server.createobject("adodb.recordset")      
            sql="select userwealth,username from [user] where usergroupid=3"      
            rs.open sql,conn,1,3      
            do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="银行行长"      
					rs2("incept")=rs(1)      
					rs2("title")="社区银行给您奖励"+cstr(money)+"元"      
					rs2("content")="社区很行给于您"+cstr(money)+"元奖励。奖励理由："+liyou      
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop      

             elseif classid=1 then      
                set rs=server.createobject("adodb.recordset")      
				sql="select userwealth,username from [user]"      
				rs.open sql,conn,1,3      
				do while not rs.eof      
					rs(0)=rs(0)+money      
					rs.update      
						set rs2=server.createobject("adodb.recordset")      
						sql2="select * from [message]"      
						rs2.open sql2,conn,1,3      
						rs2.addnew      
						rs2("sender")="银行行长"      
						rs2("incept")=rs(1)      
						rs2("title")="社区银行给您奖励"+cstr(money)+"元"      
						rs2("content")="社区很行给于您"+cstr(money)+"元奖励。奖励理由："+liyou      
						rs2("flag")=0      
						rs2("issend")=1      
						rs2.update      
						rs2.close      
					rs.movenext      
            	loop
         elseif classid=23 then       '奖励登录20次以上的会员
                     set rs=server.createobject("adodb.recordset")      
            sql="select userwealth,username from [user] where logins>20"      
            rs.open sql,conn,1,3      
            do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="银行行长"      
					rs2("incept")=rs(1)      
					rs2("title")="社区银行给您奖励"+cstr(money)+"元"      
					rs2("content")="社区很行给于您"+cstr(money)+"元奖励。奖励理由："+liyou       
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop  
         elseif classid=24 then       '奖励VIP会员
            set rs=server.createobject("adodb.recordset")      
            sql="select userwealth,username from [user] where vip=1"      
            rs.open sql,conn,1,3      
            do while not rs.eof      
                rs(0)=rs(0)+money       
                rs.update      
					set rs2=server.createobject("adodb.recordset")      
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")="银行行长"      
					rs2("incept")=rs(1)      
					rs2("title")="社区银行给您奖励"+cstr(money)+"元"      
					rs2("content")="社区很行给于您"+cstr(money)+"元奖励。奖励理由："+liyou       
					rs2("flag")=0      
					rs2("issend")=1      
					rs2.update      
					rs2.close      
                rs.movenext      
            loop  
             end if      
		else      
			
        	classuser=checkStr(trim(request.form("username")))
        	set rs=server.createobject("adodb.recordset")      
     		sql="select userwealth,username from [user] where username='"&classuser&"'"      
     		rs.open sql,conn,1,3 
			if rs.bof and rs.eof then
				Errmsg=Errmsg+"<br>"+"<li>论坛上没有["&classuser&"]这个用户"
				call bank_err()
				rs.close
				exit sub
			end if					     
        	rs(0)=rs(0)+money       
        	rs.update      
				set rs2=server.createobject("adodb.recordset")      
				sql2="select * from [message]"      
				rs2.open sql2,conn,1,3      
				rs2.addnew      
				rs2("sender")="银行行长"      
				rs2("incept")=rs(1)      
				rs2("title")="社区银行给您奖励"+cstr(money)+"元"      
				rs2("content")="社区很行给于您"+cstr(money)+"元奖励。奖励理由："+liyou  
				rs2("flag")=0      
				rs2("issend")=1      
				rs2.update      
				rs2.close      
end if      
rs.close      
     set rs=server.createobject("adodb.recordset")      
     sql="select * from [bbsnews]"      
     rs.open sql,conn,1,3      
        rs.addnew      
        rs("boardid")=0      
     	if  classid=20 then      
        	rs("title")="银行奖励所有管理员"  
			content="奖励所有管理员"    
     	elseif classid=19 then      
        	rs("title")="银行奖励所有贵宾" 
			content="奖励所有贵宾"      
        elseif classid=18 then      
        	rs("title")="银行奖励所有超级版主"  
			content="奖励所有超级版主"       
        elseif classid=17 then      
        	rs("title")="银行奖励所有版主"      
			content="奖励所有版主"   
        elseif classid=1 then      
        	rs("title")="银行奖励所有会员"      
			content="奖励所有会员"   
        else
			classuser=checkStr(trim(request.form("username")))
        	rs("title")="银行奖励用户 "&classuser     
			content="奖励用户" &classuser		
		end if      
        rs("content")="奖励理由："+liyou      
        rs("username")="银行行长"      
        rs.update      
        rs.close
		
		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content=content&" " &money&" 元，奖励理由："+liyou
			call logs("管理","银行管理中心",membername)
			sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
		end if
				  
	   sucmsg=sucmsg+"<br>"+"<li>您已经完成了奖励会员的操作，请返回"
	   call bank_suc()       
end sub 

sub savelogsetting() 
'我来了 2002.12.01   保存事件设置
'logstate logfuwu logrecordcount logadmin loguse logminmoney  loguselook
	if not master then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您没有进入银行基本设置的权限。"      
	else
		if request.form("logrecordcount")="" or (not isnumeric(request.form("logrecordcount"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>保存最新事件条数必须为数字"
		elseif request.form("logrecordcount")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>保存最新事件条数不能为负数"			
		end if	 
		if request.form("logminmoney")="" or (not isnumeric(request.form("logminmoney"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>记录储户事件(银行事件)的限制金额必须为数字" 
		elseif request.form("logminmoney")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>记录储户事件(银行事件)的限制金额不能为负数"		
		end if	
	end if	
	if founderr then      
		call bank_err()
	else
		dim log_settings
		log_settings=request.Form("logstate") & "," & request.Form("logfuwu")& "," & request.Form("logrecordcount")& "," & request.Form("logadmin")& "," & request.Form("loguse")& "," & request.Form("logminmoney")& "," & request.Form("loguselook")
		conn1.execute("update bankconfig set log_setting='"&log_settings&"'")
		
		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content="修改事件设置"
			call logs("管理","银行管理中心",membername)
			sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
		end if
				
		sucmsg=sucmsg+"<br>"+"<li>银行事件设置完成，请返回进行其他操作"
		call bank_suc()		  
	end if      
end sub

'----------------------------浏览银行事件------------------------------------------------
sub banklog()
'我来了 2002.12.01   查看事件记录
	if (not master) and cint(log_setting(6))=0 then   
		Errmsg=Errmsg+"<br>"+"<li>您没有浏览银行事件记录的权限，请与管理员联系"   
		call bank_err()      
		exit sub
	end if
	
	if clng(log_setting(2))<>0 then        '如果指定保留最新的N条事件，则删除多余的事件记录
		conn1.execute("delete from log where id not in (select top "&clng(log_setting(2))&" id from log order by id desc)")
	end if
		      
	if request("action")="dellog" then
		call batch()
	else
		call logeven()
	end if
	if founderr then call bank_err()
end sub
sub logeven()
'我来了 2002.12.01   显示事件		
	dim endpage
	dim totalrec
	dim n
	dim currentpage,page_count,Pcount
		
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if	
	if master then
		response.write "<div align=center>管理员请点击操作时间切换到管理状态</div>"
		response.write "<div align=center>管理状态下点击操作者查看该操作者的所有事件</div>"
	else
		response.write "<div align=center>浏览银行操作事件记录</div>"
	end if
	response.write "<div align=center>点击具体的类型可以浏览相同类型的操作事件记录</div>"
%>	
	<form action=Z_bank.asp?menu=13&action=dellog method=post name=even>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%"> 
		<tr>
			<th width=5% height=25>类型</th>
			<th width=15%>标题</th>
			<th width=50% id=tabletitlelink>事件内容(<a href=Z_bank.asp?menu=13 title=点击查看全部事件>查看全部事件</a>)</th>
			<th width=20% id=tabletitlelink><a href=Z_bank.asp?menu=13&action=batch&page=<%=currentpage%> title=点击切换到管理状态>操作时间</a></th>
			<th width=10%>操作人</th> 
		</tr>
<%
	set rs=server.createobject("adodb.recordset")
	if request("reaction")="银行" then
		sql="select * from log where class='银行' order by DateAndTime desc"
	elseif request("reaction")="服务" then
		sql="select * from log where class='服务' order by DateAndTime desc"
	elseif request("reaction")="管理" then	
		sql="select * from log where class='管理' order by DateAndTime desc"
	elseif request("reaction")="操作者" and trim(request("name"))<>"" then	
		sql="select * from log where UserName='"&checkStr(trim(request("name")))&"' order by DateAndTime desc"		
	else		
		sql="select * from log order by DateAndTime desc" 
	end if	
	'response.Write(sql)
	rs.open sql,conn1,1,1
	if rs.bof and rs.eof then
		response.write "<tr><td class=tablebody1 colspan=5 height=25>暂时没有任何事件</td></tr></table><br>"
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = rs.PageSize)
			response.write "<tr>"
			response.write "<td class=tablebody1 align=center height=24><a href=Z_bank.asp?menu=13&reaction="&rs("class")&" title=""点击查看所有["&rs("class")&"]事件"">"&rs("class")&"</a></td>"
			response.write "<td class=tablebody1>"&htmlencode(rs("title"))&"</td>"
			response.write "<td class=tablebody1>"&rs("content")&"</td>"
			response.write "<td class=tablebody1>"
			if request("action")="batch" and master then
				response.write "<input type=checkbox name=lid value="&rs("id")&">"
			end if
			response.write rs("DateAndTime")
			response.write "</td>"
			response.write "<td align=center class=tablebody1>"
			if master then
				if request("action")="batch" then
					response.write "<a href=Z_bank.asp?menu=13&reaction=操作者&name="&htmlencode(rs("UserName"))&" title=""操作者IP："&rs("IP")&"[点击查看该操作者的所有操作记录]"">"
				else
					response.write "<a href=dispuser.asp?name="&htmlencode(rs("UserName"))&" target=_blank title=""操作者IP："&rs("IP")&"[点击查看操作者的资料]"">"			
				end if
			else
				response.write "<a href=dispuser.asp?name="&htmlencode(rs("UserName"))&" target=_blank>"
			end if
			response.write htmlencode(rs("UserName"))&"</a></td>"	
			response.write "</tr>"
			page_count = page_count + 1
			rs.movenext
		wend
								
		if request("action")="batch" and master then
			response.write "<tr><td class=tablebody2 colspan=5>请选择要删除的事件，<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">全选 <input type=submit name=Submit value=执行  onclick=""{if(confirm('您确定执行此操作吗?')){this.document.even.submit();return true;}return false;}""></td></tr>"
		end if
		response.write "</table>"
		
		if totalrec mod Forum_Setting(11)=0 then
				Pcount= totalrec \ Forum_Setting(11)
		else
				Pcount= totalrec \ Forum_Setting(11)+1
		end if
		response.write "<table border=0 cellpadding=0 cellspacing=3 width=""97%"" align=center>"
		response.write "<tr><td valign=middle nowrap>"
		response.write "页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"
		response.write "&nbsp;每页<b>"&Forum_Setting(11)&"</b> 总数<b>"&totalrec&"</b></td>"
		response.write "<td valign=middle nowrap align=right>分页："
		if request("reaction")="操作者" then
			call DispPageNum(currentpage,PCount,"""?menu=13&reaction=操作者&name="&request("name")&"&action="&request("action")&"&page=","""")
		else
			call DispPageNum(currentpage,PCount,"""?menu=13&action="&request("action")&"&page=","""")
		end if
		response.write "</td></tr></table>"
	end if	
	rs.close
	set rs=nothing		
end sub
sub batch()
	dim lid
	if not founduser then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请登录后进行操作。"
	end if
	if not master then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>您不是系统管理员，不能管理所有日志。"
	end if

	if request.form("lid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定相关事件。"
	else
		lid=replace(request.Form("lid"),"'","")
	end if
	if founderr then exit sub
	conn1.execute("delete from log where id in ("&lid&")")
	
	if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
		content="删除指定事件"
		call logs("管理","银行事件管理",membername)
		sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
	end if	
	
	sucmsg="<li>删除指定事件成功"
	call bank_suc()
end sub
%>