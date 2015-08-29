<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="Z_bankconn.asp"-->
<!--#include file="Z_bankconfig.asp"-->
<%
dim addtili,addjinyan,addfatie,addpower,fagonggao,faduangxun,kanip,delmail
dim menu
dim fuwu_setting         '我来了添加     服务设置    2002.11.30

stats="电子银行 特殊服务"
call nav()
call head_var(0,0,"电子银行","Z_bank.asp")

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入社区服务的权限，请先登录或者同管理员联系。"
	call dvbbs_error()
elseif cint(bank_setting(6))=1 and (not master) and (DatePart("h", time)<cint(Chen_BusinessTimeSlice(0)) or DatePart("h", time)>=cint(Chen_BusinessTimeSlice(1))) then '银行是否设置定时营业
	call BankClose()	
else
    set rs=server.createobject("adodb.recordset")
    sql="select * from [fuwu]"
    rs.open sql,conn1,1,3
    addtili=rs("addtili") 			'购买魅力的价格
    addjinyan=rs("addjinyan")		'购买经验的价格
    addfatie=rs("addfatie")			'购买发帖数的价格
    addpower=rs("addpower")			'购买威望的价格
    fagonggao=rs("fagonggao")		'发公告的价格
    faduangxun=rs("faduangxun")		'发短信息的价格
    kanip=rs("kanip")				'查IP的价格
	delmail=rs("delmail")			'删除1邮件的环保基金单价   我来了添加 2002.11.30
	fuwu_setting=split(rs("fuwu_setting"),",")        '我来了添加     是否开通服务设置    2002.11.30
	                                                  'canCP/canEP/canPower/canfatie/canGonggao/canDuangxun/canKanip/canDelmail  1-是 0－否
    rs.close

	menu=request.Form("menu")	
%>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<tr><th valign=middle colspan=2 align=center height=25>电子银行  特殊服务</th></tr>
	<tr><td valign=middle class=tablebody2 height=100>
<%
	call bankhead()
	if request("action")="admin" then
		call admin()		'服务设置页面
	elseif menu="" then
		call main()       	'服务页面
	elseif menu=1 then
		call jy()			'购买经验值
	elseif menu=2 then
		call ml()			'购买魅力值
	elseif menu=3 then
		call ft()			'购买发帖数
	elseif menu=4 then
		call power()		'购买威望     
	elseif menu=5 then
		call fgg()			'发公告		
	elseif menu=6 then
		call fdx()			'群发短信
	elseif menu=7 then
		call kip()			'查会员IP
	elseif menu=8 then
		call syj()			'删除邮件
	elseif menu=10 then
		call admin1()		'保存服务设置
	end if
   
	response.write "</td></tr></table>"
end if

call activeonline()
call footer()

'----------------------------------修改价格-----------------------------------
sub admin1()
	if not master then
		Errmsg=Errmsg+"<br>"+"<li>你没有进入服务中心管理的权限!"
		call fuwu_err()
	else
		if request.form("tili")="" or (not isnumeric(request.form("tili"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>魅力价格必须为数字"
		elseif request.form("tili")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>魅力价格不能为负数"
		end if	 
		if request.form("jinyan")="" or (not isnumeric(request.form("jinyan"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>经验价格必须为数字" 
		elseif request.form("jinyan")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>经验价格不能为负数"		
		end if	
		if request.form("fatie")="" or (not isnumeric(request.form("fatie"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>发帖数价格必须为数字" 
		elseif request.form("fatie")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>发帖数价格不能为负数"			
		end if	
		if request.form("power")="" or (not isnumeric(request.form("power"))) then			
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>威望价格必须为数字" 	
		elseif request.form("power")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>威望价格不能为负数"		
		end if
		if request.form("gonggao")="" or (not isnumeric(request.form("gonggao"))) then			
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>公告价格必须为数字" 	
		elseif request.form("gonggao")<=0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>公告价格不能为负数"		
		end if
		if request.form("duangxun")="" or (not isnumeric(request.form("duangxun"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>群发短信息价格必须为数字"
		elseif request.form("duangxun")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>群发短信息价格不能为负数"		
		end if
		if request.form("kanip")="" or (not isnumeric(request.form("kanip"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>查在线会员IP信息价格必须为数字"
		elseif request.form("kanip")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>查在线会员IP信息价格不能为负数"		
		end if
		if request.form("delmail")="" or (not isnumeric(request.form("delmail"))) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>删除邮件单价必须为数字"
		elseif request.form("delmail")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>删除邮件单价不能为负数"		
		end if				
	end if	
	if founderr then      
		call fuwu_err()			
	else
		set rs=server.createobject("adodb.recordset")
		sql="select * from [fuwu]"
		rs.open sql,conn1,1,3
		rs("addtili")=request.form("tili")
		rs("addjinyan")=request.form("jinyan")
		rs("addfatie")=request.form("fatie")
		rs("addpower")=request.form("power")
		rs("fagonggao")=request.form("gonggao")
		rs("faduangxun")=request.form("duangxun")
		rs("kanip")=request.form("kanip")
		rs("delmail")=request.form("delmail")         
		
		'canCP/canEP/canPower/canfatie/canGonggao/canDuangxun/canKanip/canDelmail  1-是 0－否
		rs("fuwu_setting")=request.Form("canCP") & "," & request.Form("canEP")& "," & request.Form("canPower")& "," & request.Form("canfatie")& "," & request.Form("canGonggao")& "," & request.Form("canDuangxun")& "," & request.Form("canKanip")& "," & request.Form("canDelmail")
		
		rs.update
		rs.close

		if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
			content="调整服务价格"
			call logs("服务","服务管理中心",membername)
			sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
		end if
		sucmsg=sucmsg+"<br>"+"<li>设置完成，请返回!"		
		call fuwu_suc()
	end if	
end sub

'-------------------------------删除邮件-----------------------------------
sub syj()
	if cint(fuwu_setting(7))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>该服务已经暂停，请与管理员联系"
		call fuwu_err()
		exit sub
	end if
	founderr=false
	set rs=server.createobject("adodb.recordset")
	sql="select * from [message] where flag=1 and issend=1 and incept='"&membername&"' "
	rs.open sql,conn,1,3
	
	if not (rs.eof and rs.bof) then       '是否有已读的短信息
		rs.close
		dim cunm    '我来了添加 统计短信息条数 2002.11.30
		set rs=conn.execute("select count(incept) from [message] where flag=1 and issend=1 and sender<>'"&membername&"' and incept='"&membername&"'")
		cunm=cint(rs(0))		'统计已读的又不是自己发给自己的短信息 2002.11.30
		rs.close
		
		'我来了修改  删除所有已读短信息 2002.11.30
	  	conn.execute("delete from [message] where flag=1 and issend=1 and incept='"&membername&"'")
		
		'我来了添加 删除自己发给自己的信件不发环保奖金 2002.11.30
		if cunm=0 then
			Errmsg=Errmsg+"<br>"+"<li>你的收件箱只有自己发给自己的信件，所以您没有得到环保奖金"
			call fuwu_err()			
			exit sub		
		end if
			
		sql="select * from [user] where username='"&membername&"' "
		rs.open sql,conn,1,3
		rs("userwealth") = rs("userwealth")+delmail*cunm
		rs.update
		rs.close   
	 
		sql = "select * from [bankconfig]"
		rs.open sql,conn1,1,3            
		rs("chubei")=rs("chubei")-delmail*cunm
		rs.update
		rs.close
		set rs=nothing
		
		if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
			content="清空已读收件，得到环保奖金：<font color=red>"&delmail&"元/封×"&cunm&"封="&delmail*cunm&"元</font>"
			call logs("服务","删除短信息",membername)
			sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
		end if
		sucmsg=sucmsg+"<br>"+"<li>你的收件箱已经清空，得到环保奖金<font color=red>"&delmail&"元/封×"&cunm&"封="&delmail*cunm&"元</font>"		
		call fuwu_suc()		
	else
		rs.close
		set rs=nothing
		Errmsg=Errmsg+"<br>"+"<li>您的收件箱中没有已读短信息!"
		call fuwu_err()		
	end if
 
end sub

'------------------------------看IP--------------------------------
sub kip()
	if cint(fuwu_setting(6))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>该服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub
	end if
	
	dim rsip,sqlip
	set rs=server.createobject("adodb.recordset")
	set rsip=server.createobject("adodb.recordset")
	sql="select userwealth from [user] where username='"&membername&"' "
	rs.open sql,conn,1,3
	if request.form("user")="" then
		Errmsg=Errmsg+"<br>"+"<li>你要看谁的IP呀？"
		call fuwu_err()		
	elseif rs(0)<kanip then
		Errmsg=Errmsg+"<br>"+"<li>你的现金不够!"
		call fuwu_err()		
	elseif trim(request.form("user"))="我来了" then
		Errmsg=Errmsg+"<br>"+"<li><font color=blue>我来了说：</font><font color=red>拜托，不要查我的IP，好吧，这次免费的喔，感谢您对我的关心!</font>"
		call fuwu_err()		
	else
		
		sqlip="select IP from [online] where username='"&checkStr(trim(request.form("user")))&"' "
		rsip.open sqlip,conn,1,3
		if rsip.eof then
			Errmsg=Errmsg+"<br>"+"<li>对不起此用户不在线"
			call fuwu_err()			
		else
%>    
<table border=0 cellspacing=1 cellpadding=3 align=center class=tableborder1><tr><th height=26 >查看 <font color=red><%=htmlencode(request.form("user"))%></font> 的IP信息</th></tr><tr><td align=center height=70 class=tablebody1>会员：<font color=navy><%=htmlencode(request.form("user"))%></font> 的IP是：<font color=red><%=rsip(0)%></font><br><br>来自：<font color=red><%=address(rsip(0))%></font><br></td></tr><tr><td align=center height=26 class=tablebody2><a href="Z_fuwu.asp">返回</a></td></tr></table>          
<%          
			rsip.close          
			rs("userwealth")=rs("userwealth")-kanip          
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+kanip          
			rs.update          
			rs.close 
			  
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="查看 <font color=navy>"&htmlencode(request.form("user"))&"</font> 的IP信息，支付服务费用："&kanip&" 元"
				call logs("服务","查看在线会员IP",membername)
			end if
		       
		end if          
	end if          
          
end sub         
            
'-------------------------------群发短信------------------------------------          
sub fdx()   
	if cint(fuwu_setting(5))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>该服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub
	end if
	       
	dim rsdx,sqldx          
	set rs=server.createobject("adodb.recordset")          
	set rsdx=server.createobject("adodb.recordset")          
	sql="select userwealth from [user] where username='"&membername&"' "          
	rs.open sql,conn,1,3          
	if request.form("title")="" then          
		Errmsg=Errmsg+"<br>"+"<li>标题不能为空,请输入标题"
		call fuwu_err()            
	elseif request.form("content")="" then          
		Errmsg=Errmsg+"<br>"+"<li>你的内容是空的,必须填写内容"
		call fuwu_err()              
	elseif rs("userwealth")<faduangxun then          
		Errmsg=Errmsg+"<br>"+"<li>你的现金不够!"
		call fuwu_err()     
	else          
		rs("userwealth")=rs("userwealth")-faduangxun          
		rs.update          
		rs.close          
        if request.form("stype") = 1 then 
			content="短消息群发对象：所有在线会员"		         
			sql="select * from [online]"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
				sqldx="select * from [message]"          
				rsdx.open sqldx,conn,1,3          
				rsdx.addnew          
				rsdx("sender")=membername          
				rsdx("incept")=rs("username")          
				rsdx("title")=request.form("title")          
				rsdx("content")=request.form("content")          
				rsdx("issend")=1          
				rsdx.update          
				rsdx.close          
				rs.movenext          
			loop          
         elseif request.form("stype") = 2 then          
		 	content="短消息群发对象：所有贵宾"		 
			sql="select * from [user] where usergroupid=8"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
				sqldx="select * from [message]"          
				rsdx.open sqldx,conn,1,3          
				rsdx.addnew          
				rsdx("sender")=membername          
				rsdx("incept")=rs("username")          
				rsdx("title")=request.form("title")          
				rsdx("content")=request.form("content")          
				rsdx("issend")=1          
				rsdx.update          
				rsdx.close          
				rs.movenext          
			loop          
         elseif request.form("stype") = 3 then    
		 	content="短消息群发对象：所有超级版主和版主"		       
			sql="select * from [user] where usergroupid=3 or usergroupid=2"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
				sqldx="select * from [message]"          
				rsdx.open sqldx,conn,1,3          
				rsdx.addnew          
				rsdx("sender")=membername          
				rsdx("incept")=rs("username")          
				rsdx("title")=request.form("title")          
				rsdx("content")=request.form("content")          
				rsdx("issend")=1          
				rsdx.update          
				rsdx.close          
				rs.movenext          
			loop          
		elseif request.form("stype") = 4 then          
		 	content="短消息群发对象：所有管理员"
			sql="select * from [user] where usergroupid=1"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
				sqldx="select * from [message]"          
				rsdx.open sqldx,conn,1,3          
				rsdx.addnew          
				rsdx("sender")=membername          
				rsdx("incept")=rs("username")          
				rsdx("title")=request.form("title")          
				rsdx("content")=request.form("content")          
				rsdx("issend")=1          
				rsdx.update          
				rsdx.close          
				rs.movenext          
			loop          
		elseif request.form("stype") = 5 then 
			content="短消息群发对象：所有 管理员、超级版主、版主、贵宾"		 		         
			sql="select * from [user] where usergroupid=1 or usergroupid=2 or usergroupid=3 or usergroupid=8"          
			rs.open sql,conn,1,3          
			do while not rs.eof          
			sqldx="select * from [message]"          
			rsdx.open sqldx,conn,1,3          
			rsdx.addnew          
			rsdx("sender")=membername          
			rsdx("incept")=rs("username")          
			rsdx("title")=request.form("title")          
			rsdx("content")=request.form("content")          
			rsdx("issend")=1          
			rsdx.update          
			rsdx.close          
			rs.movenext          
			loop          
		end if          
		rs.close          
		sql = "select * from [bankconfig]"          
		rs.open sql,conn1,1,3                      
		rs("chubei")=rs("chubei")+faduangxun          
		rs.update          
		rs.close
		     
		if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
			content=content&"，支付服务费用："&faduangxun&" 元"
			call logs("服务","短消息群发",membername)
			sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
		end if
		sucmsg=sucmsg+"<br>"+"<li>短消息群发成功，请返回"		
		call fuwu_suc()		       
	end if          
end sub           
       
'---------------------------------发公告=-------------------------------          
sub fgg()    
	if cint(fuwu_setting(4))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>该服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub
	end if
	      
	if trim(request.form("title"))="" then
		founderr=true 
		Errmsg=Errmsg+"<br>"+"<li>公告标题不能为空,请输入标题" 
	end if
	if trim(request.form("content"))="" then
		founderr=true 
		Errmsg=Errmsg+"<br>"+"<li>公告内容是空的,必须填写内容" 
	end if
	if mymoney<fagonggao then
		founderr=true 
		Errmsg=Errmsg+"<br>"+"<li>你的现金不够" 
	end if
	BoardID=int(request.form("boardid"))		
         
	if founderr then
		call bank_err()        
	else          
		conn.execute("update [user] set userwealth=userwealth-"&fagonggao&" where userid="&userid)
		set rs=server.createobject("adodb.recordset")		         
		sql="select * from [bbsnews]"          
		rs.open sql,conn,1,3          
		rs.addnew          
		rs("boardid")=BoardID        
		rs("username")=request.form("username")          
		rs("title")=request.form("title")          
		rs("content")=request.form("content")          
		rs.update          
		rs.close 
		
		myCache.name="AnnounceMents"&BoardID
		myCache.makeEmpty	
		         
		sql = "select * from [bankconfig]"          
		rs.open sql,conn1,1,3                      
		rs("chubei")=rs("chubei")+fagonggao          
		rs.update          
		rs.close
		
		if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
			content="发布论坛公告，支付服务费用："&fagonggao&" 元"
			call logs("服务","发布公告",membername)
			sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
		end if
		sucmsg=sucmsg+"<br>"+"<li>恭喜你，公告发布成功" 		
		call fuwu_suc()		
	end if          
end sub           
          
'--------------------------------购买威望值------------------------- 我来了 添加         
sub power() 
	if cint(fuwu_setting(2))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>该服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub
	end if
	 
	dim buypower  
	buypower=0      
  	set rs=server.createobject("adodb.recordset")          
  	sql="select userwealth,userpower from [user] where username='"&membername&"' "          
  	rs.open sql,conn,1,3          

		if request.form("power")="" then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>请输入您要购买的威望点数"
		elseif not isnumeric(request.form("power")) then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>购买的威望点数必须是数字"
		elseif request.form("power")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>请不要输入负数"		
		else
			buypower=abs(int(request.form("power")))
			if rs(0)<buypower*addpower then         
				founderr=true 
				Errmsg=Errmsg+"<br>"+"<li>您的现金不够"
			end if			
		end if
			
		if founderr then
			call bank_err()		 
      	else          
			rs(0)=rs(0)-buypower*addpower          
			rs(1)=rs(1)+buypower          
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+addpower          
			rs.update          
			rs.close 
			
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="购买威望值，支付服务费用："&buypower&"点 × "&addpower&" 元/点 = "& buypower*addpower & " 元"
				call logs("服务","购买威望值",membername)
				sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
			end if
			sucmsg=sucmsg+"<br>"+"<li>购买威望值成功，花费："&buypower&"点 × "&addpower&" 元/点 ="& buypower*addpower & "元"		
			call fuwu_suc()			         
      end if           
end sub           
          
'--------------------------------购买发帖数-------------------------          
sub ft() 
	if cint(fuwu_setting(3))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>该服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub
	end if
	         
	dim fatie 
	fatie=0         
	set rs=server.createobject("adodb.recordset")          
	sql="select userwealth,Article from [user] where username='"&membername&"' "          
	rs.open sql,conn,1,3  
          
		if request.form("fatie")="" then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>请输入您要购买的帖子数"
		elseif not isnumeric(request.form("fatie")) then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>购买的帖子数必须是数字"
		elseif request.form("fatie")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>请不要输入负数"		
		else
			fatie=abs(int(request.form("fatie")))
			if rs(0)<fatie*addfatie then         
				founderr=true 
				Errmsg=Errmsg+"<br>"+"<li>您的现金不够"
			end if			
		end if
			
		if founderr then
			call bank_err()		 
		else          
			rs(0)=rs(0)-fatie*addfatie          
			rs(1)=rs(1)+fatie          
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+fatie*addfatie          
			rs.update          
			rs.close   
			
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="购买发帖数，支付服务费用："&fatie&"点 × "&addfatie&" 元/点 = "& fatie*addfatie & " 元"
				call logs("服务","购买发帖数",membername)
				sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
			end if
			sucmsg=sucmsg+"<br>"+"<li>购买发帖数成功，花费："&fatie&"点 × "&addfatie&" 元/点 ="& fatie*addfatie & "元"		
			call fuwu_suc()			        
		end if          
end sub 
          
'--------------------------------购买魅力值-------------------------          
sub ml()  
	if cint(fuwu_setting(0))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>该服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub
	end if
	        
	dim tili   
	tili=0       
	set rs=server.createobject("adodb.recordset")          
	sql="select userwealth,usercp from [user] where username='"&membername&"' "          
	rs.open sql,conn,1,3 
		if request.form("tili")="" then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>请输入您要购买的魅力值"
		elseif not isnumeric(request.form("tili")) then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>购买的魅力值必须是数字"
		elseif request.form("tili")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>请不要输入负数"		
		else
			tili=abs(int(request.form("tili")))
		end if
      	if rs(0)<tili*addtili then         
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>您的现金不够"
		end if			
		if founderr then
			call bank_err()           
      	else     
			rs(0)=rs(0)-tili*addtili          
			rs(1)=rs(1)+tili 		     
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+tili*addtili          
			rs.update          
			rs.close  
			 
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="购买魅力值，支付服务费用："&tili&"点 × "&addtili&" 元/点 = "& tili*addtili & " 元"
				call logs("服务","购买魅力值",membername)
				sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
			end if
			sucmsg=sucmsg+"<br>"+"<li>购买魅力值成功，花费："&tili&"点 × "&addtili&" 元/点 ="& tili*addtili & "元"		
			call fuwu_suc()			       
      	end if           
end sub 
          
'--------------------------------购买经验值-------------------------          
sub jy() 
	if cint(fuwu_setting(1))=0 then 
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>该服务已经暂停，请与管理员联系"
		call bank_err()
		exit sub
	end if
	         
	dim jinyan     
	jinyan=0     
	set rs=server.createobject("adodb.recordset")          
	sql="select userwealth,userep from [user] where username='"&membername&"' "          
	rs.open sql,conn,1,3
       
		if request.form("jinyan")="" then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>请输入您要购买的经验值"
		elseif not isnumeric(request.form("jinyan")) then
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>购买的经验值必须是数字"
		elseif request.form("jinyan")<0 then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>请不要输入负数"		
		else
			jinyan=abs(int(request.form("jinyan")))
		end if
      	if rs(0)<jinyan*addjinyan then         
			founderr=true 
			Errmsg=Errmsg+"<br>"+"<li>您的现金不够"
		end if			
		if founderr then
			call bank_err()        
        else          
			rs(0)=rs(0)-jinyan*addjinyan          
			rs(1)=rs(1)+jinyan          
			rs.update          
			rs.close          
			sql = "select * from [bankconfig]"          
			rs.open sql,conn1,1,3                      
			rs("chubei")=rs("chubei")+jinyan*addjinyan          
			rs.update          
			rs.close
			
			if cint(log_setting(0))=1 and cint(log_setting(1))=1 then
				content="购买经验值，支付服务费用："&jinyan&"点 × "&addjinyan&" 元/点 = "& jinyan*addjinyan & " 元"
				call logs("服务","购买经验值",membername)
				sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
			end if
			sucmsg=sucmsg+"<br>"+"<li>购买经验值成功，花费："&jinyan&"点 × "&addjinyan&" 元/点 ="& jinyan*addjinyan & "元"		
			call fuwu_suc()			       
      end if           
end sub 
        
'-------------------------------------------主程序-----------------------------------          
sub main()    
'fuwu_setting : canCP/canEP/canPower/canfatie/canGonggao/canDuangxun/canKanip/canDelmail  1-是 0－否  我来了 2002.11.30
'fuwu_setting():  0      1      2        3         4           5         6         7      
%>          
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">          
<tr><th align=center height=26><b>社区服务中心</b></th></tr>          
<tr><td align=center height=26 class=tablebody1>欢迎来到社区服务中心,你现有的现金<font color="#FF0000"><%=mymoney%></font>元。我们将给您提供优质的服务,请选择您的服务项目吧!</td></tr>
<tr><td align=left height=26 class=tablebody2> <b>特殊服务</b>　<%if cint(fuwu_setting(0))=1 or cint(fuwu_setting(1))=1 or cint(fuwu_setting(2))=1 or cint(fuwu_setting(3))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%></td></tr>       
          
<form name="form1" method="post" action="Z_fuwu.asp"><tr ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(1))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%>　　　　<b>购买经验值</b> 请输入您要购买多少经验值:<INPUT name=jinyan> <INPUT type=submit value=购买 name=submit> 单价:<font color="#FF0000"><%=addjinyan%></font>元 　　　　我的经验值: <font color="#FF0000"><%=myuserep%></font></td></tr><INPUT type=hidden value="1" name=menu></form>          
          
<form name="form2" method="post" action="Z_fuwu.asp"><tr ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(0))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%>　　　　<b>购买魅力值</b> 请输入您要购买多少魅力值:<INPUT name=tili> <INPUT type=submit value=购买 name=submit class=tablebody1> 单价:<font color="#FF0000"><%=addtili%></font>元 　　　　我的魅力值: <font color="#FF0000"><%=myusercp%></font></td></tr><INPUT type=hidden value="2" name=menu></form>          
          
<form name="form3" method="post" action="Z_fuwu.asp"><tr ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(3))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%>　　　　<b>购买发帖数</b> 请输入您要购买多少发帖数:<INPUT name=fatie> <INPUT type=submit value=购买 name=submit> 单价:<font color="#FF0000"><%=addfatie%></font>元 　　　　我的发帖数: <font color="#FF0000"><%=myArticle%></font></td></tr><INPUT type=hidden value="3" name=menu></form>          
          
<form name="form4" method="post" action="Z_fuwu.asp"><tr ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(2))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%>　　　　<b>购买<font color="#FF0000">威望值</font></b> 请输入您要购买多少威望值:<INPUT name=power> <INPUT type=submit value=购买 name=submit> 单价:<font color="#FF0000"><%=addpower%></font>元 　　　　我的威望值: <font color="#FF0000"><%=mypower%></font></td></tr><INPUT type=hidden value="4" name=menu></form>          
          
<tr ><td align=left height=30 class=tablebody2> <b>公告发布</b> 公告发布每次需要<font color="#FF0000"><%=fagonggao%></font>元 　<%if cint(fuwu_setting(4))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%></td></tr>          
<form name="form5" method="post" action="Z_fuwu.asp">
<tr><td align=center height=30 class=tablebody1> 
                 
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:70%">          
<tr>          
<th colspan="2">公告发布</th>          
</tr>
<tr>          
<td width="25%" class=tablebody2>发布版面:</td>          
<td align=left class=tablebody1>          
<% set rs=server.createobject("adodb.recordset")         
sql="select boardid,boardtype from board"          
   rs.open sql,conn,1,1          
%>          
        <select name="boardid" size="1">          
          <option value="0">论坛首页</option>
          <%
	do while not rs.eof
        response.write "<option value='"+CStr(rs("BoardID"))+"'>"+rs("Boardtype")+"</option>"+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
%>
        </select>          
</td></tr>          
<tr>           
      <td width="25%" valign=top class=tablebody2>发布人：</td>          
      <td width="75%" class=tablebody1>          
        <input type=text name=username size=36 value="<%=membername%>" disabled>          
        <input type=hidden name=username value="<%=membername%>">          
      </td>          
    </tr>          
    <tr>           
      <td width="25%" valign=top class=tablebody2>标题：</td>          
      <td width="75%" class=tablebody1>          
        <input type=text name=title size=36>          
      </td>          
    </tr>          
    <tr>           
      <td width="25%" valign=top class=tablebody2>内容：</td>          
      <td width="75%" class=tablebody1>          
        <textarea cols=70 rows=6 name="content"></textarea>          
      </td>          
    </tr>          
    <tr>          
      <td width="100%" valign=top colspan="2" align=center class=tablebody2><input type=Submit value="发 送" name=Submit>&nbsp;<input type="reset" name="Clear" value="清 除"></td>          
    </tr>          
</table>          
          
</td></tr>
<INPUT type=hidden value="5" name=menu>
</form>          
          
<tr><td align=left height=30 class=tablebody2> <b>短消息群发</b> 每次<font color=red><%=faduangxun%></font>元　<%if cint(fuwu_setting(5))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%></td></tr>          
          
          
<form name="form6" method="post" action="Z_fuwu.asp"><tr  ><td align=center height=30 class=tablebody1>            
<table cellspacing=1 cellpadding=3 class=tableborder1 style="width:70%">          
<tr>          
<th colspan="2">短消息群发</th>          
</tr>
<tr>           
                  <td width="30%" class=tablebody2>消息标题</td>          
                  <td width="70%" class=tablebody1>           
                    <input type="text" name="title" size="50">          
                  </td>          
                </tr>          
<tr>           
                  <td width="30%" class=tablebody2>接收方选择</td>          
                  <td width="70%" class=tablebody1>           
                    <select name=stype size=1>          
					<option value="1">所有在线用户</option>
					<option value="2">所有贵宾</option>
					<option value="3">所有版主</option>
					<option value="4">所有管理员</option>
					<option value="5">贵宾/版主/管理员</option>
					</select>
                  </td>
                </tr>
                <tr> 
                  <td width="30%" height="20" valign="top" class=tablebody2>
                    <p>消息内容</p>
                  </td>
                  <td width="70%" height="20" class=tablebody1> 
                    <textarea name="content" cols="70" rows="10"></textarea>
                  </td>
                </tr>
                <tr> 
                  <td colspan=2 align=center height="23" class=tablebody2><input type="submit" name="Submit" value="发送消息">  <input type="reset" name="Submit2" value="重新填写">
                  </td>
                </tr>
</table>
</td></tr>
<INPUT type=hidden value="6" name=menu>
</form>
<tr  height=30 ><td align=left height=26 class=tablebody2> <b>其他服务</b>　<%if cint(fuwu_setting(7))=1 or cint(fuwu_setting(6))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%></td></tr>
<form name="form7" method="post" action="Z_fuwu.asp">
<tr  ><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(6))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%>　　　　<b>查看在线会员IP</b>  每次<font color=red><%=kanip%></font>元 查看 <INPUT name=user> 的IP和来源 <INPUT type=submit value=查看 name=submit></td></tr>          
<INPUT type=hidden value="7" name=menu>
</form>          
          
<form name="form8" method="post" action="Z_fuwu.asp">          
<tr><td align=left height=30 class=tablebody1><%if cint(fuwu_setting(7))=1 then%><font color=blue>√</font><%else%><font color=red>×</font><%end if%>　　　　<b>删除已读短消息</b>  每次删除将得到银行 <font color=red><%=delmail%>元×短信息条数 </font> 环保奖金  <INPUT type=submit value=删除 name="b1" onclick="javascript:{if(confirm('您确定执行删除操作吗?')){return true;}return false;}"> (<font color="#FF0000">注意：此功能将删除你的所有已读邮件，不可恢复！</font>)</td></tr>          
<INPUT type=hidden value="8" name=menu>
</form>          
</table>
<br>          
<%          
end sub          

sub fuwu_err() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>服务错误信息</th></tr>
	<tr><td width="100%" class=tablebody1><b>产生错误的可能原因：</b><br><br><li>您是否仔细阅读了<a href="boardhelp.asp?boardid=<%=boardid%>">帮助文件</a>，可能您还没有登录或者不具有使用当前功能的权限
	<font color=<%=Forum_body(8)%>><%=errmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href="javascript:history.go(-1)">返回上一页</a></td></tr>
</table>
<%
end sub         

sub fuwu_suc() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>服务操作成功</th></tr>
	<tr><td width="100%" class=tablebody1><b>操作成功：</b><br><br><li>欢迎光临<%=Forum_info(0)%> 社区服务中心
	<font color=navy><%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>返回上一页</a></td></tr>
</table>
<%
end sub

 '-----------------------------变量设置---------------------------------          
sub admin() 
if not master then
	Errmsg=Errmsg+"<br>"+"<li>你没有进入服务中心管理的权限,请与管理员联系"
	call fuwu_err() 	
else	
'fuwu_setting : canCP/canEP/canPower/canfatie/canGonggao/canDuangxun/canKanip/canDelmail  1-是 0－否  我来了 2002.11.30
'fuwu_setting():  0      1      2        3         4           5         6         7
%>          
	<br>
	<table border=0 cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">          
	<tr  height=22><th colspan=3  height=26 >服务中心管理----服务价格设置</th></tr>          
	<form name="form1" method="post" action="Z_fuwu.asp">           
	<tr >          
		<td align=right width="35%" class=tablebody1>每点经验的价格：</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=jinyan value="<%=addjinyan%>"> 元</td> 
		<td align=left width="30%" class=tablebody1>是否开通该服务 
		<input type=radio name="canEP" value=0 <%if cint(fuwu_setting(1))=0 then%>checked<%end if%>>否&nbsp;
		<input type=radio name="canEP" value=1 <%if cint(fuwu_setting(1))=1 then%>checked<%end if%>>是&nbsp;
		</td>	        
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1>每点魅力的价格：</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=tili value="<%=addtili%>"> 元</td>          
		<td align=left width="30%" class=tablebody1>是否开通该服务 
		<input type=radio name="canCP" value=0 <%if cint(fuwu_setting(0))=0 then%>checked<%end if%>>否&nbsp;
		<input type=radio name="canCP" value=1 <%if cint(fuwu_setting(0))=1 then%>checked<%end if%>>是&nbsp;
		</td>	
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1>每点发帖的价格：</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=fatie value="<%=addfatie%>">  元</td>
		<td align=left width="30%" class=tablebody1>是否开通该服务 
		<input type=radio name="canfatie" value=0 <%if cint(fuwu_setting(3))=0 then%>checked<%end if%>>否&nbsp;
		<input type=radio name="canfatie" value=1 <%if cint(fuwu_setting(3))=1 then%>checked<%end if%>>是&nbsp;
		</td>	         
	</tr>         
	<tr >         
		<td align=right width="35%" class=tablebody1>每点威望的价格：</td>        
		<td align=left width="35%" class=tablebody1><INPUT name=power value="<%=addpower%>">  元</td>        
		<td align=left width="30%" class=tablebody1>是否开通该服务 
		<input type=radio name="canPower" value=0 <%if cint(fuwu_setting(2))=0 then%>checked<%end if%>>否&nbsp;
		<input type=radio name="canPower" value=1 <%if cint(fuwu_setting(2))=1 then%>checked<%end if%>>是&nbsp;
		</td>	
	</tr>        
	<tr >        
		<td align=right width="35%" class=tablebody1>发公告的价格：</td>        
		<td align=left width="35%" class=tablebody1><INPUT name=gonggao value="<%=fagonggao%>"> 元</td>          
		<td align=left width="30%" class=tablebody1>是否开通该服务 
		<input type=radio name="canGonggao" value=0 <%if cint(fuwu_setting(4))=0 then%>checked<%end if%>>否&nbsp;
		<input type=radio name="canGonggao" value=1 <%if cint(fuwu_setting(4))=1 then%>checked<%end if%>>是&nbsp;
		</td>	
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1>群发短信的价格：</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=duangxun value="<%=faduangxun%>"> 元</td>  
		<td align=left width="30%" class=tablebody1>是否开通该服务 
		<input type=radio name="canDuangxun" value=0 <%if cint(fuwu_setting(5))=0 then%>checked<%end if%>>否&nbsp;
		<input type=radio name="canDuangxun" value=1 <%if cint(fuwu_setting(5))=1 then%>checked<%end if%>>是&nbsp;
		</td>	        
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1>看IP的价格：</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=kanip value="<%=kanip%>"> 元</td> 
		<td align=left width="30%" class=tablebody1>是否开通该服务 
		<input type=radio name="canKanip" value=0 <%if cint(fuwu_setting(6))=0 then%>checked<%end if%>>否&nbsp;
		<input type=radio name="canKanip" value=1 <%if cint(fuwu_setting(6))=1 then%>checked<%end if%>>是&nbsp;
		</td>	         
	</tr>          
	<tr >          
		<td align=right width="35%" class=tablebody1> 删除<font color=red>1</font>封邮件单价：</td>          
		<td align=left width="35%" class=tablebody1><INPUT name=delmail value="<%=Delmail%>"> 元</td>  
		<td align=left width="30%" class=tablebody1>是否开通该服务 
		<input type=radio name="canDelmail" value=0 <%if cint(fuwu_setting(7))=0 then%>checked<%end if%>>否&nbsp;
		<input type=radio name="canDelmail" value=1 <%if cint(fuwu_setting(7))=1 then%>checked<%end if%>>是&nbsp;
		</td>	         
	</tr>          
	<tr> 	       
		<td align=center colspan=3 class=tablebody2><INPUT type=submit value=修改 name=submit></td>          
	</tr> 
	<INPUT type=hidden value="10" name=menu>         
	</form> 	         
	</table> 
         
	<% 
	end if
end sub 
%>