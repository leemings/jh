<!--#include file="z_gp_Function.asp"-->
<%Dim ChengJiaoNum,ChengJiaoMoney,Gupiao_Setting
Dim Trade_Setting
Dim StopGPReadme
Dim AI_Setting
Dim Custom_Setting		'今日成交笔数,今日成交金额,股票设置
Dim MyUserID,MyCash,MyBiShu,MyStats		'用户ID,可用的资金,本日买卖笔数,状态  
Dim HaveAccount '是否有股票账号 
Dim errmess,sucmess,rUrl
	
if membername="" then
	response.write "<title>股票交易所-提示信息</title><style type=text/css>td {  font-family: 宋体; font-size: 9pt}</style>"
	errmess="<li>您还没有登录，请登录后再来！"
	call endinfo(0)
	response.end
end if
	
REM 读取设置信息 
set rs=gp_conn.execute("select top 1 * from [GupiaoConfig] order by id") 
if datediff("d",rs("TodayDate"),Now())<>0 then
	gp_conn.execute("update [KeHu] set JinRiMaiRu=0,JinRiMaiChu=0")
	gp_conn.execute("update [GupiaoConfig] set TodayBuy=0,TodaySale=0,TodayTotal=0,TodayDate=date()")  '
	call History()
	ChengJiaoNum=0
	ChengJiaoMoney=0
else
	ChengJiaoNum=rs("TodayBuy")+rs("TodaySale")
	ChengJiaoMoney=rs("TodayTotal")
end if
StopGPReadme=rs("StopReadme")
Gupiao_Setting=split(rs("Gupiao_Setting"),",")
Gupiao_Setting(0)=CLNG(Gupiao_Setting(0))
Gupiao_Setting(1)=CLNG(Gupiao_Setting(1))
Gupiao_Setting(2)=CLNG(Gupiao_Setting(2))
Gupiao_Setting(3)=CLNG(Gupiao_Setting(3))
Gupiao_Setting(6)=CLNG(Gupiao_Setting(6))
Gupiao_Setting(7)=CLNG(Gupiao_Setting(7))
Gupiao_Setting(9)=CLNG(Gupiao_Setting(9))
Gupiao_Setting(10)=CSNG(Gupiao_Setting(10))
Gupiao_Setting(11)=CSNG(Gupiao_Setting(11))
Gupiao_Setting(12)=CLNG(Gupiao_Setting(12))
Trade_Setting=split(rs("Trade_Setting"),"|")
Trade_Setting(0)=CLNG(Trade_Setting(0))
Trade_Setting(2)=CLNG(Trade_Setting(2))
Trade_Setting(3)=CLNG(Trade_Setting(3))
Trade_Setting(4)=CLNG(Trade_Setting(4))
Trade_Setting(5)=CSNG(Trade_Setting(5))
Trade_Setting(6)=CLNG(Trade_Setting(6))
Trade_Setting(7)=CLNG(Trade_Setting(7))
Trade_Setting(8)=CLNG(Trade_Setting(8))
Trade_Setting(9)=CSNG(Trade_Setting(9))
Trade_Setting(10)=CSNG(Trade_Setting(10))
Trade_Setting(11)=CSNG(Trade_Setting(11))
Trade_Setting(12)=CSNG(Trade_Setting(12))
Trade_Setting(13)=CSNG(Trade_Setting(13))
Trade_Setting(14)=CSNG(Trade_Setting(14))
Trade_Setting(15)=CSNG(Trade_Setting(15))
if ubound(Trade_Setting)<16 then
	redim preserve Trade_Setting(16)
	Trade_Setting(16)=1
else
	Trade_Setting(16)=CLNG(Trade_Setting(16))
end if
if ubound(Trade_Setting)<17 then
	redim preserve Trade_Setting(17)
	Trade_Setting(17)=0
else
	Trade_Setting(17)=CLNG(Trade_Setting(17))
end if
AI_Setting=split(rs("AI_Setting"),"|")
AI_Setting(0)=CLNG(AI_Setting(0))
AI_Setting(1)=CSNG(AI_Setting(1))
Custom_Setting=split(rs("Custom_Setting"),"||")
rs.close

if not master then
	if Gupiao_Setting(0)=0 then
		stats="交易大厅"
		call nav()
		call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
		call CloseGuPiao(0)
		response.end	
	elseif Gupiao_Setting(3)=1 then  				'股票 是否 收市 
		if hour(time)<CLNG(split(Gupiao_Setting(4),"||")(0)) or hour(time)>=CLNG(split(Gupiao_Setting(4),"||")(1)) then
			stats="交易大厅"
			call nav()
			call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
			call CloseGuPiao(1)
			response.end
		end if		
	end if
end if
 
if Gupiao_Setting(6)=1 and (not master) then
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
	if (not isnull(session("GP_ReflashTime"))) and Gupiao_Setting(7)>0 and DoReflashPage then
		if DateDiff("s",session("GP_ReflashTime"),Now())<Gupiao_Setting(7) then
			stats="交易大厅"
			call nav()
			call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
			call CloseGuPiao(2)
			response.end
		else
			session("GP_ReflashTime")=Now()
		end if
	elseif isnull(session("GP_ReflashTime")) and Gupiao_Setting(7)>0 and DoReflashPage then
		Session("GP_ReflashTime")=Now()
	end if
end if
		
REM 读取用户信息
set rs=gp_conn.execute("select SuoDing,ZiJin,JinRiMaiRu,JinRiMaiChu,id,ZongZiJin from [KeHu] where ZhangHao='"&membername&"'"  )
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
elseif rs(0)=2 then
	HaveAccount=false
	MyCash=rs(1)
	MyBiShu=rs(2)+rs(3)
	MyUserID=rs(4)
else
	HaveAccount=True	'是否已经开户 
	MyCash=rs(1)		'可用的资金 
	MyBiShu=rs(2)+rs(3) '今日买卖笔数 
	MyUserID=rs(4)		'UserID 
end if
rs.close	
%>
