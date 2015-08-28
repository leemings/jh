<!--#include file="Function.asp"-->
<%
	Dim ChengJiaoNum,ChengJiaoMoney,Gupiao_Setting,Trade_Setting,StopReadme,AI_Setting,Custom_Setting		'今日成交笔数,今日成交金额,股票设置
	Dim membername,master,userhidden
	Dim MyUserID,MyCash,MyBiShu,MyStats		'用户ID,可用的资金,本日买卖笔数,状态  
	Dim HaveAccount '是否有股票账号 
	Dim errmess,sucmess,rUrl,founderr,i,stats
		
	
	membername=session("yx8_mhjh_username")
	userhidden=""
	userhidden=iif(userhidden="",2,userhidden)
	master=false 
	
	if membername="" then
		response.write "<title>股票交易所-提示信息</title><style type=text/css>td {  font-family: 宋体; font-size: 9pt}</style>"
		errmess="<li>您还没有登陆，请登陆后再来！"
		call endinfo(0)
		response.end
	end if
	
'	set rsbbs=connbbs.execute("select id from [admin] where adduser='"&membername&"'")
'	if rsbbs.eof and rsbbs.bof then
'		master=false
'	else
'		master=true
'	end if
'	rsbbs.close
        if Session("yx8_mhjh_grade")=10 and instr(Application("yx8_mhjh_admin"),Session("yx8_mhjh_username"))<>0 then
           master=true
        else
	   master=false        
        end if 
			
	REM 读取设置信息 
	set rs=conn.execute("select top 1 * from [GupiaoConfig] order by id") 
	'if datediff("d",rs("TodayDate"),Now())<>0 then
	'过整点就股票开盘
	if datediff("h",rs("TodayDate"),Now())<>0 then
		conn.execute("update [客户] set 今日买入=0,今日卖出=0")
		conn.execute("update [GupiaoConfig] set TodayBuy=0,TodaySale=0,TodayTotal=0,TodayDate=now()")  '
		call History()
		ChengJiaoNum=0
		ChengJiaoMoney=0
	else
		ChengJiaoNum=rs("TodayBuy")+rs("TodaySale")
		ChengJiaoMoney=rs("TodayTotal")
	end if
	StopReadme=rs("StopReadme")
	Gupiao_Setting=split(rs("Gupiao_Setting"),",")
	Trade_Setting =split(rs("Trade_Setting"),"|")
	AI_Setting    =split(rs("AI_Setting"),"|")
	Custom_Setting=split(rs("Custom_Setting"),"||")
	rs.close
	
	REM 是否限时开收股市/是否暂时关闭股市
	if not master then
		if cint(Gupiao_Setting(0))=0 then
			call CloseGuPiao(0)
			response.end	
		elseif cint(Gupiao_Setting(3))=1 then  				'股票 是否 收市 
			if hour(time)<cint(split(Gupiao_Setting(4),"||")(0)) or hour(time)>=cint(split(Gupiao_Setting(4),"||")(1)) then
				call CloseGuPiao(1)
				response.end
			end if		
		end if
	end if
	 
	REM 防刷新功能 
	if cint(Gupiao_Setting(6))=1 and (not master) then
		Dim SplitReflashPage
		Dim DoReflashPage
		Dim ScriptName
		DoReflashPage=false
		ScriptName=lcase(request.ServerVariables("PATH_INFO"))
		if trim(Gupiao_Setting(8))<>"" then
			SplitReflashPage=split(Gupiao_Setting(8),"|")
			for i=0 to ubound(SplitReflashPage)
				if instr(scriptname,SplitReflashPage(i))>0 then
					DoReflashPage=true
					exit for
				end if
			next
		end if
		if (not isnull(session("GP_ReflashTime"))) and cint(Gupiao_Setting(7))>0 and DoReflashPage then
			if DateDiff("s",session("GP_ReflashTime"),Now())<cint(Gupiao_Setting(7)) then
				'response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT="&Gupiao_Setting(7)&">本页面起用了防刷新机制，请不要在<font color=red>"&Gupiao_Setting(7)&"</font>秒内连续刷新本页面<BR><font color=blue>"&Gupiao_Setting(7)&"</font>秒之后将会自动打开页面，请稍后……"
				call CloseGuPiao(2)
				response.end
			else
				session("GP_ReflashTime")=Now()
			end if
		elseif isnull(session("GP_ReflashTime")) and cint(Gupiao_Setting(7))>0 and DoReflashPage then
			Session("GP_ReflashTime")=Now()
		end if
	end if
			
	REM 读取用户信息
	set rs=conn.execute("select 锁定,资金,今日买入,今日卖出,id,总资金 from [客户] where 帐号='"&membername&"'"  )
	if rs.eof and rs.bof then 
		HaveAccount=false
		MyCash=0
		MyBiShu=0
		MyUserID=0
	elseif rs(0)=1 then 	
		rs.close
		response.write "<title>"&Gupiao_Setting(5)&"-账号被冻结</title><style type=text/css>td {  font-family: 宋体; font-size: 9pt}</style>"
		errmess="<li>您的股票账户被冻结，请与管理员联系！"
		call endinfo(0)
		response.end
	else
		HaveAccount=True	'是否已经开户 
		MyCash=rs(1)		'可用的资金 
		MyBiShu=rs(2)+rs(3) '今日买卖笔数 
		MyUserID=rs(4)		'UserID 
	end if
	rs.close	
%>