<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_ViewManage_Conn.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
dim page
if request("action")="reset" and master and membername<>"" then
	connvm.execute("delete * from [fuser]")
	Response.Redirect "z_ViewManage.asp"
end if
If Request("page") = "" or not isInteger(Request("page")) then
	page = 1
Else
	page = CINT(Request("page"))
End If
if not founduser then
	stats="版主评估"
	call nav()
	call head_var(2,0,"","")
	Errmsg=Errmsg+"<br>"+"<li>您没有查看版主评估的权限，请先登录或者同管理员联系。"
	call dvbbs_error()
else
	if request("action")="hy" then
		stats="会员情况"
		call nav()
		call head_var(2,0,"","")
		call hy()
	else
		stats="版主评估"
		call nav()
		call head_var(2,0,"","")
		call bz()	
	end if 
end if
call activeonline()
call footer() 

sub bz()
	dim Gzdays,Gzlbt,Gzhmon,Gzwealth,Gzuserep,Gzusercp,Gzpower,Gzifmess,Gzjbgz,pagenum,ifAnn,ifAutoCancelMSG
	dim gzifarticle,gzsuperfj,gzparticle,gzplogins,gzplogs,gziflogs,gziflogins,ifautocancel,autocday,gzmasterfj
	Gzdays=7		'工资日间隔，每7天发放一次
	gzifmess="1"		'发薪是否短信通知版主  1是，0否
	pagenum=8		'每页工资条列出多少行
	
	Gzjbgz=100		'版主的每版基本工资
	gzsuperfj=500		'超级版主附加工资
	gzmasterfj=1000		'超级版主附加工资
	ifautocancel="1"	'是否自动取消失职版主职务  1自动，0手动
	ifAnn="1"		'是否发布公告  1是，0否
	ifAutoCancelMSG="1"		'是否向已被取消资格的版主发出短信通知  1是，0否
	autocday=15		'超过多少天未登录取消版主职务
	
	gzifarticle="1"		'积分计算公式是否计算发帖量，1计算，0不计算
	gzparticle=0.5		'发帖量计算系数，如100发帖，计入工作积分50
	gziflogins="1"		'积分计算公式是否计算登录次数，1计算，0不计算
	gzplogins=0.5		'登录次数计算系数，如登录100*0.5，实际计入50
	gziflogs="1"		'积分计算公式是否计算维护量，1计算，0不计算
	gzplogs=1		'维护量计算系数，默认为1
	
	Gzlbt=20		'得到业绩工资所需工作积分底线
	gzhmon=20		'积分每增加多少涨一级业绩工资
	gzwealth=100		'一级业绩工资加多少钱
	gzuserep=50		'一级业绩工资加多少经验
	gzusercp=20		'一级业绩工资加多少魅力
	gzpower=1		'一级业绩工资加多少威望
  
	'------------按照你的实际需求修改以上参数-----
	dim board,rsu,UserName,Absentday,article,Message,RsB,Bcount,BoardName,allbm,RsL,RsF,Rstmp,p
	dim Bgzday,Btgz,savemoney,saveep,savecp,savepower
	dim Bnumgz,Bnumlog,saveMessage,masterlist,Barticle
	response.write "<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>"
	response.write "<tr>"
	response.write "<th valign=middle colspan=7 align=center height=25>版主工作情况</th>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td colspan=7 class=tablebody1><br>"
	response.write "　　感谢各位版主对 <font color="&forum_body(8)&">"&Forum_info(0)&"</font> 的大力支持，"
	response.write "本站每 <fotn color="&forum_body(8)&">"&gzdays&"</font> 天对各位版主进行积分奖励。<br><br>"
	response.write "　　工作积分的计算公式为："
	response.write "版主 <b>"&gzdays&"</b> 天内"
	if gziflogins="1" then 
    response.write "登录次数 <b>*</b> <b>"&formatnumber(gzplogins,1,-1)&"</b>"
    if gzifarticle="1" or gziflogs="1" then response.write " <b>+</b> "
	end if
	if gzifarticle="1" then 
		response.write "本期发帖量 <b>*</b> <b>"&formatnumber(gzparticle,1,-1)&"</b>"
		if gziflogs="1" then response.write " <b>+</b> "
	end if
	if gziflogs="1" then response.write "版面维护次数 <b>*</b> <b>"&formatnumber(gzplogs,1,-1)&"</b>"
	response.write "。<br>"
	response.write "　　实际给付工资为基本工资加上业绩工资，维护次数为零的版主将被停薪。<br>"
	response.write "　　版主的基本工资为每版 <font color="&forum_body(8)&"><b>"&Gzjbgz&"</b></font> 元，超级版主另有 <font color="&forum_body(8)&"><b>"&gzsuperfj&"</b></font> 元附加奖励，管理员另有 <font color="&forum_body(8)&"><b>"&gzmasterfj&"</b></font> 元附加奖励。<br>"
	response.write "　　工作积分超过 <font color="&forum_body(8)&"><b>"&Gzlbt&"</b></font> 点后给付业绩工资，业绩工资计算方法：<br>"
	response.write "　　版主的工作积分每提高 <font color="&forum_body(8)&"><b>"&gzhmon&"</b></font> 点，"
	response.write "奖励现金 <font color="&forum_body(8)&"><b>"&gzwealth&"</b></font> 元，"
	response.write "威望 <font color="&forum_body(8)&"><b>"&gzpower&"</b></font> 点，"
	response.write "经验 <font color="&forum_body(8)&"><b>"&gzuserep&"</b></font> 点，"
	response.write "魅力 <font color="&forum_body(8)&"><b>"&gzusercp&"</b></font> 点。<br>"
	response.write "　　请各位版主注意，超过 <font color="&forum_body(8)&"><b>"&autocday&"</b></font> 天未登录的版主将被自动取消版主资格。<br><br>"
	response.write "　　感谢大家的支持！<br><br>"
	response.write "</td>"
	response.write "</tr>"
	response.write "<tr height=25>"
	response.write "<th><B>版主</b></th>"	
  response.write "<th><B>担任版块</b></th>"
	response.write "<th><B>发表文章数</b></th>"
	response.write "<th><B>发薪/登录日</b></th>"
	response.write "<th><B>登录</b></th>"
	response.write "<th><B>维护量/发帖</b></th>"
	response.write "<th><B>预期工资/评语</b></th>"
	response.write "</tr>"
	sql="select UserName,UserGroupID,Article,lastlogin,logins from [user] where UserGroupID < 4 order by lastlogin desc"
	set RsU=Server.createobject("ADODB.Recordset")
	RsU.open sql,conn,1,1
	if rsu.eof or rsu.bof then
		response.write "<tr><td class=tablebody1 align=center colspan=7 height=30>论坛中没有任何版主！</td></tr>"
	else
		Rsu.PageSize = Pagenum
		if page<1 then page=1
		If page>RsU.Pagecount Then page=RsU.Pagecount              
		rsu.AbsolutePage=page
		set RSB=Server.createobject("ADODB.Recordset")
		allbm=""
		do while not RsU.eof
			UserName=RsU("UserName")
			if isnull(RsU("lastlogin")) then
				AbsentDay=autocday+1
			else
				AbsentDay=DateDiff("d",RsU("lastlogin"),Now)
			end if
			article=RsU("article")
			if AbsentDay =< autocday/5  then
				Message="<font face=Wingdings color=blue>v</font> <font color=#0080d5>工作比较认真负责！</font>"
			elseif AbsentDay =< autocday*2/5  then
				Message="已松散，但 <font color="&forum_body(8)&"><b>"&int(autocday*2/5)&"</b></font> 天内有登录！"
			elseif AbsentDay =< autocday*3/5 then
				Message="注意，已经 <font color="&forum_body(8)&"><b>"&int(autocday*2/5)&"</b></font> 天以上没有登录了！</font>"
			elseif AbsentDay =< autocday*4/5 then
				Message="工作失职，已经 <font color="&forum_body(8)&"><b>"&int(autocday*3/5)&"</b></font> 天以上没有登录了！</font>"
			elseif AbsentDay =< autocday-1 then
				Message="严重失职，已经 <font color="&forum_body(8)&"><b>"&int(autocday*4/5)&"</b></font> 天以上没有登录了，将考虑撤消职务！</font>"
			else
				Message="<font color="&forum_body(8)&"><b>这个版主不称职，已经被撤消职务！</b></font>"
				if allbm="" then
					allbm=UserName
				else
					allbm=allbm&"，"&UserName
				end if
		    if ifautocancel="1" then
					connvm.execute("delete from fuser where username='"&UserName&"'")
					dim rs5,BoardMaster_1,BoardMaster_2,k
					dim oldMinArticle
					oldMinArticle=99999999
					set rs5=conn.execute("select * from usertitle order by MinArticle desc")
					do while not rs5.eof
						conn.execute("update [user] set usergroupid=4,userclass='"&rs5("usertitle")&"',titlepic='"&rs5("titlepic")&"' where username='"&UserName&"' and (article<"&oldMinArticle&" and article>="&rs5("MinArticle")&" )")
						oldMinArticle=rs5("MinArticle")
						rs5.movenext
					loop
					rs5.close
					set rs5=conn.execute("select BoardMaster,boardid from board where instr('|'&BoardMaster&'|','|"&UserName&"|')>0")
					do while not rs5.eof
						BoardMaster_2=""
						k=0
						BoardMaster_1=split(rs5(0),"|") 
						for i=0 to ubound(BoardMaster_1)
							if lcase(trim(UserName))=lcase(trim(BoardMaster_1(i))) then 
								BoardMaster_1(i)="" 
							end if
							if trim(BoardMaster_1(i))<>"" then
								if BoardMaster_2="" then
									BoardMaster_2=BoardMaster_1(i)
								else
									BoardMaster_2=BoardMaster_2&"|"&trim(BoardMaster_1(i))
								end if
							end if
						next
						conn.execute("update [board] set BoardMaster='"&BoardMaster_2&"' where boardid="&rs5(1)&"")
						rs5.movenext
					loop
					set rs5=nothing
					if ifAutoCancelMSG then
						conn.execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&Username&"','"&forum_info(0)&"','系统消息：您的版主资格已取消','原因：超过"&autocday&"天没有登录系统，工作严重失职，情节恶劣！',Now(),0,1)")
					end if
				end if
			end if
			BCount=0
			set RsB=conn.execute("select BoardType,BoardMaster,BoardID from [board] where instr('|'&BoardMaster&'|','|"&UserName&"|')>0")
			if not RsB.eof then
				BoardName=""
				do while not RsB.eof
					if BoardName="" then
						BoardName="<a href=list.asp?boardid="&RsB("BoardID")&" target=_blank>"&RsB("BoardType")&"</a>"
					else
						BoardName=BoardName&"<br>"&"<a href=list.asp?boardid="&RsB("BoardID")&" target=_blank>"&RsB("BoardType")&"</a>"
					end if
					Bcount=Bcount+1
					RsB.movenext
				loop
			else 
				BoardName="<font color="&forum_body(8)&"><b>没有担任的版块</b></font>"
			end if
			set RsF=connvm.execute("select * from [fuser] where username='"&Username&"' ")
			if RsF.eof then
				set Rstmp=server.createobject("adodb.recordset")
				Rstmp.open "select count(l_announceid) from [log] where l_username='"&UserName&"' group by l_announceid",conn,1,1
				connvm.execute("insert into fuser(username,FuLiDate,oldlogs,oldlogins,oldarticle) values('"&UserName&"',date()-1,"&Rstmp.recordcount&","&RsU("logins")&","&article&")")
				set Rstmp=nothing
			end if
			set Rsl=server.createobject("adodb.recordset")
			Rsl.open "select count(l_announceid) from [log] where l_username='"&UserName&"' group by l_announceid",conn,1,1
			if RsF.eof then
				Bnumgz=0
				Bnumlog=0
				Barticle=0
				Bgzday=(date()-1+gzdays)
			else
				Bnumgz=cint(Rsl.recordcount-Rsf("oldlogs"))
				Bnumlog=cint(RsU("logins")-Rsf("oldlogins"))
				Bgzday=(RsF("Fulidate")+gzdays)
				Barticle=(RsU("article")-Rsf("oldarticle"))
			end if
			'积分计算公式
			if gziflogins<>"1" then gzplogins=0
			if gzifarticle<>"1" then gzparticle=0
			if gziflogs<>"1" then gzplogs=0
			Btgz=(Bnumlog*gzplogins+Barticle*gzparticle+Bnumgz*gzplogs)
			if AbsentDay<=gzdays and Bnumgz>0 then
				if Btgz>Gzlbt then
					savemoney=(((Btgz-Gzlbt)\Gzhmon)*Gzwealth+gzjbgz*Bcount)
					savepower=(((Btgz-Gzlbt)\Gzhmon)*Gzpower)
					saveep=(((Btgz-Gzlbt)\Gzhmon)*Gzuserep)
					savecp=(((Btgz-Gzlbt)\Gzhmon)*Gzusercp)
				elseif Btgz>=0 then
					savemoney=gzjbgz*Bcount
					savepower=0
					saveep=0
					savecp=0
				end if
				if savemoney>0 and RsU("usergroupid")=2 then savemoney=cint(savemoney+gzsuperfj)
				if savemoney>0 and RsU("usergroupid")=1 then savemoney=cint(savemoney+gzmasterfj)
			else
			  savemoney=0
			  savepower=0
			  saveep=0
			  savecp=0
			end if
			if date()>=Bgzday and not RsF.eof and savemoney>0 then
			  savemessage="("&FormatDateTime(Rsf("Fulidate"),2)&"至"&FormatDateTime(Bgzday,2)&")：工作积分("&cint(Btgz)&")，金钱("&savemoney&")，威望("&savepower&")，经验("&saveep&")，魅力("&savecp&")"
			  connvm.execute("update [fuser] set Fulidate=date(),oldlogs="&Rsl.recordcount&",oldlogins="&RsU("logins")&",lastinfo='"&savemessage&"',oldarticle="&Rsu("article")&" where username='"&Username&"'")
			  conn.execute("update [user] set userwealth=userwealth+"&savemoney&",userep=userep+"&saveep&",usercp=usercp+"&savecp&",userpower=userpower+"&savepower&" where username='"&Username&"'")
			  if gzifmess="1" then
			    savemessage=savemessage&"。　再次对您为本站所做工作表示感谢！继续努力呦:)"
					conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&Username&"','"&Forum_info(0)&"','系统消息：您的工资','"&checkSTR(savemessage)&"',Now(),0,1)")
				end if
			end if
			response.write "<tr>"
			response.write "<td rowspan=2 class=tablebody1 align=left>&nbsp;"
			response.write "<a href=javascript:openScript('messanger.asp?action=new&touser="&UserName&"',500,400)>"
			select case Rsu("usergroupid")
			case 1
				response.write "<img src="&Forum_info(7)&Forum_pic(0)&" alt=给"&UserName&"发消息 height=18 border=0>"
			case 2
				response.write "<img src="&Forum_info(7)&Forum_pic(16)&" alt=给"&UserName&"发消息 height=18 border=0>"
			case 3
				response.write "<img src="&Forum_info(7)&Forum_pic(1)&" alt=给"&UserName&"发消息 height=18 border=0>"
			end select
			response.write "</a>&nbsp;"
			response.write "<a href=dispuser.asp?name="&UserName&" target=_blank title=""查看"&htmlencode(UserName)&"的信息"">"
			if membername=username then
				response.write "<font color=blue>"
			end if
			response.write UserName
			if membername=username then
				response.write "</font>"
			end if
			response.write "</a>"
			response.write "</td>"
			response.write "<td rowspan=2 class=tablebody1 align=center>"
			response.write BoardName
			response.write "</td>"
			response.write "<td rowspan=2 class=tablebody1 align=center>文章 <font color="&forum_body(8)&"><b>"&article&"</b></font> 篇 </td>"
			response.write "<td class=tablebody1>发薪日期：<b>"&FormatDateTime(Bgzday,2)&"</b></td>"
			response.write "<td class=tablebody1>本期：<font color=blue><b>"&Bnumlog&"</b></font></td>"
			response.write "<td class=tablebody1>维护量：<font color=blue><b>"&Bnumgz&"</font></td>"
			response.write "<td class=tablebody1 height=20 valign=top><img src=pic/mybbs.gif width=16 height=16 align=absmiddle alt=该版主上期工资："
			if Rsf.eof then 
        response.write "0"
      else
        response.write Rsf("lastinfo")
      end if
			response.write ">&nbsp;<font color=blue>积分："
			response.write "<b><font color=red>"&cint(Btgz)&"</font></b> <font color=black>|</font> "
			if AbsentDay<=gzdays and Bnumgz>0 then  
				response.write "工资：金钱(<b><font color="&forum_body(8)&">"&savemoney&"</font></b>)"
			else
				response.write " <b><font color="&forum_body(8)&">此版主本期被停薪</font></b>"
			end if
			if savepower>0 then response.write "威望(<b><font color="&forum_body(8)&">"&savepower&"</font></b>)"
			if saveep>0 then response.write "经验(<b><font color="&forum_body(8)&">"&saveep&"</font></b>)"
			if savecp>0 then response.write "魅力(<b><font color="&forum_body(8)&">"&savecp&"</font></b>)"
			response.write "</font></td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td class=tablebody1>最后登录：<b>"&FormatDateTime(rsU("lastlogin"),2)&"</b></td>"
			response.write "<td  class=tablebody1>失踪：<font color=blue><b>"&AbsentDay&"</b></font></td>"
			response.write "<td class=tablebody1>发帖量：<font color=blue><b>"&Barticle&"</b></font></td>"
			response.write "<td class=tablebody1><font color=#0080d5>评定："&Message&"</font></td>"
			response.write "</tr>"
			p=p+1  
			RsF.close
			Set RsF=nothing
			RsL.close
			set Rsl=nothing
			RsU.movenext
  		if p>= rsu.PageSize then exit do
		loop
		response.write "<tr>"
		response.write "<td align=center height=25 class=tablebody2><a href=z_viewmanage.asp?action=hy><font color=blue>[会员情况]</font></a></td>"
		response.write "<td height=25 colspan=6 class=tablebody2>"
		response.write "<marquee width=90% scrolldelay=100 scrollamount=4 onmouseout=""if (document.all!=null){this.start()}"" onmouseover=""if (document.all!=null){this.stop()}"">"
		if allbm<>"" then
			response.write allbm&"已经严重失职，将被撤消版主职务！"
		else
			response.write "所有版主工作积极，<b>"&forum_info(0)&"</b>感谢大家的努力！"
		end if
		response.write "</MARQUEE>"
		response.write "</td></tr></table>"
		response.write "<table cellpadding=0 cellspacing=0 width="&forum_body(12)&" align=center>"
		response.write "<tr>"
		response.write "<tr><td height=25 valign=middle colspan=5 align=left>&nbsp;&nbsp;◇&nbsp;<b>"&Forum_info(0)&"</b> 目前共有版主 <b>"&rsu.RecordCount&"</b> 人，第 <b>"&page&"</b> / <b>"&rsu.PageCount&"</b> 页</td>"
		response.write "<td height=25 valign=middle colspan=2 align=right>"
		if master then response.write "<a href=z_viewmanage.asp?action=reset>[<font color=red><b title=""注意：只有在你清空了管理日志log表后，才需要进行此项操作"">重建工资数据</b></font>]</a>&nbsp;"
		call disppagenum(page,rsu.pagecount,"?page=","")
		response.write "</td>"
		response.write "</tr>"
	end if
	response.write "</table>"
	rsU.close
	Set RsU=nothing
	if ifAnn and not allbm="" then
		conn.execute("insert into bbsnews(boardid,title,content,username,addtime) values (0,'［公告］撤消"&allbm&" 版主一职！','原因：超过 "&autocday&" 天没拥锹枷低虫，工作严重失职，情节恶劣！','"&forum_info(0)&"',now())")
	end if
end sub

sub hy()
	dim GlTime,sex
	dim PageCount
	PageCount=10
	set rs = server.CreateObject ("adodb.recordset")
	sql="select username,adddate,Sex,Userclass,UserGroup,lastlogin from [user] where logins <= 1 and Article=0 and datediff('d',lastlogin,date())>15 order by lastlogin"
	rs.open sql,conn,1,1
	response.write "<table Class=TableBorder1 cellspacing=1 cellpadding=3 align=center>"
	response.write "<tr>"
	response.write "<th height=25 align=center colspan=6>会员登录情况</td>"
	response.write "</tr>"
	if rs.eof or rs.bof then
		response.write "<tr><td colspan=6 class=tablebody1 height=30 align=center>所有会员全部比较热心！</td></tr>"
		response.write "<tr>"
		response.write "<td class=tablebody2 align=center height=25 colspan=6><a href=z_viewmanage.asp><font color=blue>[版主工作]</font></a></td>"
		response.write "</tr>"
	else
		response.write "<tr>"
		response.write "<td colspan=6 class=tablebody1><br>"
		response.write "　　欢迎光临 <font color="&forum_body(8)&">"&Forum_info(0)&"</font> ！<br><br>"
		response.write "　　这里列出的是只登录过一次，又没发表过任何文章的会员，看到这么多的人上黑名单真让人伤心！<br>"
		response.write "　　所以如果你打算只来一次的话，请不要注册了！<br>"
		response.write "　　这样即节省了数据库的空间，又省了你的时间，不是吗？<br><br>"
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr height=25>"
		response.write "<th>会员姓名</th>"
		response.write "<th>注册日期</th>"
		response.write "<th>消失天数</th>"
		response.write "<th>会员性别</th>"
		response.write "<th>会员门派</th>"
		response.write "<th>发送信息</th>"
		response.write "</tr>"
		rs.AbsolutePage=Page
		rs.PageSize=PageCount
		do while not rs.EOF
			response.write "<tr height=25 align=center>"
			response.write "<td class=tablebody1 align=center><a href=dispuser.asp?name="&rs(0)&" target=""_blank"" title=""查看"&rs(0)&"的个人资料"">"&rs(0)&"</a></td>"
			response.write "<td class=tablebody1 align=center>"&rs(1)&"</td>"
			GlTime= datediff("d",rs(5),date())
			response.write "<td class=tablebody1 align=center><b><font color=red>"&GlTime&"</font></b> 天</td>"
			if rs(2)=1 then
				sex="男"
			elseif rs(2)=0 then
				sex="女" 
			end if
			response.write "<td class=tablebody1 align=center>"&sex&"</td>"
			response.write "<td class=tablebody1 align=center>"&rs(4)&"</td>"
			response.write "<td class=tablebody1 align=center><a target=""_blank"" href=messanger.asp?action=new&touser="&rs("username")&"><img src=pic/message.gif alt=给他留言 border=0></a></td>"
			response.write "</tr>" 
			i=i+1
			rs.movenext
			if i>= rs.PageSize then exit do
		loop  
		response.write "<tr>"
		response.write "<td class=tablebody2 align=center height=25><a href=z_viewmanage.asp><font color=blue>[版主工作]</font></a></td>"
		response.write "<td class=tablebody2>&nbsp;&nbsp;共 <b>"&rs.RecordCount&"</b> 人 第 <b>"&page&"</b> / <b>"&rs.PageCount&"</b> 页</td>"
		response.write "<td class=tablebody2 align=right colspan=5>"
		call disppagenum(page,rs.pagecount,"?page=","&action=hy")
		response.write "</td>"
		response.write "</tr>"
	end if
	response.write "</table>"
	rs.Close
	set rs = nothing 
end sub
%>
