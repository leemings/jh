<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="Z_TestConn.ASP"-->
<%
'---------------------
'edit by 绿水青山/我来了
'---------------------
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入开心辞典的权限，请先登录或者同管理员联系。"
	call dvbbs_error()
else
	stats="开心辞典 有奖知识问答"
	
	call nav()
	call head_var(0,0,"开心辞典","Z_test.asp")

	dim ars,userp,ajs,intime
	set ars=server.createobject("adodb.recordset")
	
	set ars=aconn.execute("select usersinew,js,intime from [testuser] where username='"&membername&"'")
	userp=ars(0)		'游戏币
	ajs=ars(1)			'剩余答题数
	intime=ars(2)		'上次答题时间
	ars.close

	if userp<=0 then		'游戏币用完了
		session("checkpage")=""
		session("ajj")=""
		session("atitle")=""
		session("akey1")=""
		session("akey1")=""
		session("akey2")=""
		session("akey3")=""
		session("akey4")=""
		session("aok")=""
		session("aid")=""
		session("testlb")=""
		session("atime")=""
		errmsg=errmsg+"<br>"+"<li>你的游戏币已经用完了,不能再答题了，请去发帖或上传题目挣点游戏币吧！"
		founderr=true
		call test_err()
	else
%>
 	<table align=center cellspacing=1 cellpadding="3" class=tableborder1>
		<tr><th><%=stats%></th></tr>
		<tr><td class=tablebody2>
<%	
		dim AnswerSetting
		set ars=aconn.execute("select AnswerSetting from [config]")
		AnswerSetting=split(ars(0),"||")		'0每回答一题所需游戏币个数 ||1每回答一题魅力值增减值||2每回答一题经验值增减值||3两次玩之间的时间间隔||4答题数||5是否显示正确答案||6花钱买知识所需金钱||7答题是否限时||8题目是否支持UBB
		if ubound(AnswerSetting)<8 then
			redim preserve AnswerSetting(8)
			AnswerSetting(8)="1"
		end if
		ars.close
		
		if cint(AnswerSetting(8))=1 then
			dim abgcolor,username
			%><!--#include file="inc/ubbcode.asp" --><%
		end if

		if request("action")="hd" then
		   call hd()		' 回答选择题
		elseif request("action")="dt" then
		   call dt()		' 准备出题
		elseif request("action")="out" then
		   call out()		' 退出答题
		elseif request("action")="hdtk" then
		   call hdtk()		' 回答填空题
		else
			call out()		' 退出答题
		end if
	end if
end if
response.Write "</td></tr></table>"
set ars=nothing
aconn.close
set aconn=nothing

call activeonline() 
call footer()
'=================================================
'-------------------答题－－判断---------------
sub dt()
	'response.Write("<font color=red>现在在dt()函数</font>")
	if request("btest")<>"" or session("checkpage")<>"on" then
		if AnswerSetting(3)-datediff("n",intime,now())>=1 then
			dim timec		'还差多少分钟才可以答题
			timec=AnswerSetting(3)-datediff("n",intime,now())
			session("checkpage")=""
			session("ajj")=""
			session("atitle")=""
			session("akey1")=""
			session("akey1")=""
			session("akey2")=""
			session("akey3")=""
			session("akey4")=""
			session("aok")=""
			session("aid")=""
			session("testlb")=""
			session("atx")=""
			session("acount")=""
			session("aokcount")=""
			errmsg=errmsg+"<br>"+"<li>您上次玩的时间是 ◇ "&intime&" ◇ ,离下一次玩还差 ◇ "&timec&" ◇ 分钟"
			founderr=true
			call test_err()
		else	'只在回答第一题的时候执行
			aconn.execute ("update [testuser] set js="&AnswerSetting(4)&",intime=now() where username='"&membername&"'")
			ajs=AnswerSetting(4)
			session("checkpage")="on"
			session("oklu")=0		'本次答题正确率
			session("ok")=0			'本次答题答对数			
			call dt1()	'出题
		end if
	else
		call dt1()		'出题
	end if
end sub 
'-------------------答题－－显示题目---------------
sub dt1()
	session("atime")=0
	if session("testlb")="" then
		session("testlb")=request("testlb")
	end if

	if ajs<=0 then	'本次答题数已经为零了
		response.Write "<script>alert('你已回答了"&AnswerSetting(4)&"题，请过"&AnswerSetting(3)&"分钟后再来，谢谢光临！')</script>"
		call out()		
	else
		aconn.execute ("update [testuser] set js=js-1 where username='"&membername&"'")
	end if
	
	dim ausername
	sql="select jj,time,title,key1,key2,key3,key4,ok,id,tx,tcount,okcount,username  from [test] where lb="&session("testlb")
	ars.CursorLocation=3
	ars.open sql,aconn,1,1
	Randomize
	ars.absolutePosition=Fix(ars.recordcount*Rnd)+1
	session("ajj")=ars(0)			'本题 奖金
	session("atime")=ars(1)			'本题 限时
	session("atitle")=ars(2)		'题目
	session("akey1")=ars(3)			'			
	session("akey2")=ars(4)			'
	session("akey3")=ars(5)			'
	session("akey4")=ars(6)			'
	session("aok")=ars(7)			'正确答案代号
	session("aid")=ars(8)			'题目ID
	session("tx")=ars(9)			'题目类型
	session("acount")=ars(10)		'答题人数
	session("aokcount")=ars(11)		'答对人数
	ausername=ars(12)				'题目上传者
	ars.close

	if cint(AnswerSetting(7))=1 then	'答题是否限制时间 
		%><script>window.setTimeout('GetSec()',1)</script><%
	end if
	if session("tx")=0 then	'选择题 
%>
		<form name="form1" method="post" action="z_testii.asp?action=hd">
		<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:80%">
		<tr>
			<td width="70"  align="center" class=tablebody2><b>您的属性</b></td>
			<td width="600" colspan="4" class=tablebody2>游戏币：<%=userp%>&nbsp;&nbsp;经验：<%=myuserep%>&nbsp;&nbsp;魅力：<%=myusercp%>&nbsp;&nbsp;现金：<%=mymoney%> &nbsp;&nbsp;&nbsp;您剩余的答题数：<font color=blue><%=ajs%></font></td>
		</tr>
		<tr>
			<td width="70"  align="center" class=tablebody1><b>题　　号</b></td>
			<td width="51"  class=tablebody1>&nbsp;<%=session("aid")%> </td>
			<td width="105"  class=tablebody1 align="center"><b>本题奖金</b></td>
			<td width="101"  class=tablebody1><font color="#FF0000">± <%=session("ajj")%> </font></td>
			<td width="337"  class=tablebody1><span lang="zh-cn"><%if cint(AnswerSetting(7))=1 then%><b>本题限时</b> <%=session("atime")%> 秒　&nbsp;<b>所剩时间</b>&nbsp;<input size=6 name=timespent>秒<%else%><font color=blue>&nbsp;已设置不限时</font><%end if%></span></td> 
		</tr>
		<tr>
			<td width="70"  align="center" class=tablebody2><b>题&nbsp;&nbsp;&nbsp;目</b></td> 
			<td width="600" colspan="4" align="center" class=tablebody2><table border=0 width="100%" cellspacing=1 class=tablebody2><tr><td><%if cint(AnswerSetting(8))=1 then%><%=DvBCode(session("atitle"),1,2)%><%else%><%=htmlencode(session("atitle"))%><%end if%></td></tr></table></td>
		</tr>
		<tr>
			<td width="70" height="20" align="center" class=tablebody1>
			<input type="radio" value=1 name="R1" checked></td>
			<td width="600" colspan="4" height="20" class=tablebody1>&nbsp;<%=session("akey1")%>　</td>
		</tr>
		<tr>
			<td width="70" height="20" align="center" class=tablebody1><input type="radio" value=2 name="R1"></td>
			<td width="600" colspan="4" height="20" class=tablebody1>&nbsp;<%=session("akey2")%>　</td>
		</tr>
		<%if trim(session("akey3"))<>"" then%>
		<tr>
			<td width="70" height="20" align="center" class=tablebody1><input type="radio" value=3 name="R1"></td>
			<td width="600" colspan="4" height="20" class=tablebody1>&nbsp;<%=session("akey3")%>　</td>
		</tr>
		<%end if
		if trim(session("akey4"))<>"" then%>
		<tr>
			<td width="70" height="20" align="center" class=tablebody1><input type="radio" value=4 name="R1"></td>
			<td width="600" colspan="4" height="20" class=tablebody1>&nbsp;<%=session("akey4")%>　</td>
		</tr>
		<%end if%>
		<tr>
			<td width="70" height="20" align="center" class=tablebody2><b>本题属性</b></td>
			<td width="600" colspan="4" height="20" class=tablebody2>已应答人数：<%=session("acount")%>&nbsp;&nbsp;&nbsp;&nbsp;答对人数：<%=session("aokcount")%>&nbsp;&nbsp;&nbsp;&nbsp;本题上传者：<font color=red><b><%if ausername="nobody" or ausername="" then%>未指定<%else%><%=ausername%><%end if%></b></font></td> 
		</tr>
		</table>
		
		<tr><td class=tablebody2 align="center">
		<input type="submit" value="提交" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="重置" name="B2">
		</td></tr>
		</form>
<%	else		' 填空题 	%>
		<form name="form1" method="post" action="z_testii.asp?action=hdtk">
		<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:80%">
		<tr>
			<td width="70" align="center" class=tablebody2><b>您的属性</b></td>
			<td width="600" colspan="4" class=tablebody2> 游戏币：<%=userp%>&nbsp;&nbsp;经验：<%=myuserep%>&nbsp;&nbsp;魅力：<%=myusercp%>&nbsp;&nbsp;现金：<%=mymoney%>&nbsp;&nbsp;&nbsp;你剩余的答题数：<font color=blue><%=ajs%></font></td> 
		</tr>
		<tr>
			<td width="70" align="center" class=tablebody1><b>题　　号</b></td>  
			<td width="51" class=tablebody1>&nbsp;<%=session("aid")%> </td>
			<td width="105" class=tablebody1 align="center"><b>&nbsp;</b><b>本题奖金</b></td>
			<td width="101" class=tablebody1>&nbsp;<font color="#FF0000">± <%=session("ajj")%> </font></td>
			<td width="337" class=tablebody1><span lang="zh-cn"><%if cint(AnswerSetting(7))=1 then%><b>本题限时</b> <font color="#FF0000"><%=session("atime")%></font> 秒　<b>所剩时间</b>&nbsp; <input size=6 name=timespent> 秒<%else%><font color=blue>&nbsp;已设置不限时</font><%end if%></span></td>
		</tr>
		<tr>
			<td width="70" align="center" class=tablebody2><b>题&nbsp;&nbsp;&nbsp;目</b></td> 
			<td width="600" colspan="4" align="center" class=tablebody2><table border=0 width="96%" cellspacing=1><tr><td><%if cint(AnswerSetting(8))=1 then%><%=DvBCode(session("atitle"),1,2)%><%else%><%=htmlencode(session("atitle"))%><%end if%></td></tr></table></td>
		</tr>
		<tr>
			<td width="70" align="center" class=tablebody1>您的答案：</td>
			<td width="600" colspan="4" class=tablebody1>&nbsp;<input type="text" name="testansw" size="50"></td>
		</tr>
		<tr>
			<td width="70" align="center" class=tablebody2><b>本题属性</b></td>
			<td width="600" colspan="4" class=tablebody2>已应答人数：<%=session("acount")%>&nbsp;&nbsp;&nbsp;&nbsp;答对人数：<%=session("aokcount")%>&nbsp;&nbsp;&nbsp;&nbsp;本题上传者：<font color=red><b><%if ausername="nobody" or ausername="" then%>未指定<%else%><%=ausername%><%end if%></b></font></td>
		</tr>
		</table>
			
		<tr><td class=tablebody2 align="center">
		  <input type="submit" value="提交" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="重置" name="B2">
		</td></tr>
		</form>
<%
	end if
	if cint(AnswerSetting(7))=1 then	'答题是否限制时间 
		session("ChenTimeStart")=timer()
	end if	
end sub
'---------------回答---选择题--------------
sub hd()
	if session("checkpage")="" or session("aid")="" then
		call out()
	end if
	if cint(AnswerSetting(7))=1 then	'答题是否限制时间 
		if timer()-session("ChenTimeStart")>session("atime")+5 then		'第二层判断是否回答超时，杜绝后退前进的非法答题
			response.Write "<script>alert('您的回答已经超出了本题的时间限制！')</script>"
			response.write "<script>top.location.href='?action=dt'</script>"
		end if
	end if
	
	dim cjs		'已经回答的题数
	cjs=AnswerSetting(4)-ajs	'cjs－已经回答的题数	ajs－剩余的题数

	if request("R1")=session("aok") then	'答案正确
		conn.execute ("update [user] set userWealth=userWealth+"&session("ajj")&",userCP=userCP+"&AnswerSetting(1)&",userEP=userEP+"&AnswerSetting(2)&" where username='"&membername&"'")
		aconn.execute ("update [testuser] set userans=userans+1,usersinew=usersinew-"&AnswerSetting(0)&",userok=userok+1 where username='"&membername&"'")
		aconn.execute ("update [test] set tcount=tcount+1,okcount=okcount+1 where id="&session("aid"))
%>
	<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:80%">
	<tr>
		<td align="center" class=tablebody2><b>您的属性</b></td> 
		<td colspan="3" class=tablebody2>游戏币：<%=userp-AnswerSetting(0)%><font color=green>↓</font>&nbsp;&nbsp;经验：<%=myuserep+AnswerSetting(2)%><font color=red>↑</font>&nbsp;&nbsp;魅力：<%=myusercp+AnswerSetting(1)%><font color=red>↑</font>&nbsp;&nbsp;现金：<%=mymoney+session("ajj")%><font color=red>↑</font> 
		<hr size="1">出题数：<%=AnswerSetting(4)%> 题&nbsp;&nbsp;已回答：<font color=blue><%=cjs%></font> 题&nbsp;&nbsp;答对：<font color=red><%=session("ok")+1%></font> 题&nbsp;&nbsp;正确率：
<%
		if cjs=0 then
			response.write "<font color=red>0.00%</font>"
		else
			session("ok")=session("ok")+1		'本次答题正确数
			session("oklu")=round(session("ok")*100/cjs,2)	'本次答题正确率
			response.write "<font color=red>"&session("oklu")&"%</font>"
		end if
%>
		</td>
	</tr>	
	<tr>
		<td width="70" align="center" class=tablebody1><b>奖&nbsp;&nbsp;&nbsp;金</b></td>
		<td class=tablebody1>&nbsp;现金 <font color="#FF0000">+ <%=session("ajj")%> </font>&nbsp;经验 <font color="#FF0000">+ <%=AnswerSetting(2)%> </font>&nbsp;魅力 <font color="#FF0000">+ <%=AnswerSetting(1)%> </font></td>
		<td class=tablebody1>&nbsp;<font color=blue>√&nbsp;<b>您的回答完全正确！</b></font></td>
		<td width="85" class=tablebody1 align="center"><a href="z_TestReport.asp?ID=<%=session("aid")%>" title="发现该题的答案有误，我要向管理员举报" target="_blank"><font color=blue>举报此题</font></a></td>
	</tr>
	<tr>
		<td align="center" class=tablebody2><b>题&nbsp;&nbsp;&nbsp;目</b></td>
		<td colspan="3" align="center" class=tablebody2><table border=0 width="100%" cellspacing=1><tr><td><%if cint(AnswerSetting(8))=1 then%><%=DvBCode(session("atitle"),1,2)%><%else%><%=htmlencode(session("atitle"))%><%end if%></td></tr></table></td>
	</tr>
   <tr>
		<td align="center" class=tablebody1>正确答案</td> 
		<td colspan="3" height="20" class=tablebody1>&nbsp;<%=session("akey"&session("aok"))%></td>
	</tr>
	</table>
	
	<tr><td class=tablebody2 align="center">
		<a href="?action=dt">继续下一题</a>　　<a href="?action=out">不玩了！</a>
	</td></tr>
<%  
	else		'回答错误 
		conn.execute ("update [user] set userWealth=userWealth-"&session("ajj")&",userCP=userCP-"&AnswerSetting(1)&",userEP=userEP-"&AnswerSetting(2)&" where username='"&membername&"'")
		aconn.execute ("update [testuser] set userans=userans+1,usersinew=usersinew-"&AnswerSetting(0)&" where username='"&membername&"'")
		aconn.execute ("update [test] set tcount=tcount+1 where id="&session("aid"))
%>
		<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:80%">
		<tr>
			<td align="center" class=tablebody2><b>您的属性</b></td>
			<td colspan="3" class=tablebody2>游戏币：<%=userp-AnswerSetting(0)%><font color=green>↓</font>&nbsp;&nbsp;经验：<%=myuserep-AnswerSetting(2)%><font color=blue>↓</font>&nbsp;&nbsp;魅力：<%=myusercp-AnswerSetting(1)%><font color=blue>↓</font>&nbsp;&nbsp;现金：<%=mymoney-session("ajj")%><font color=blue>↓</font>&nbsp;&nbsp;&nbsp;&nbsp;
			<hr size="1">出题数：<%=AnswerSetting(4)%> 题&nbsp;&nbsp;已回答：<font color=blue><%=cjs%></font> 题&nbsp;&nbsp;答对：<font color=red><%=session("ok")%></font> 题&nbsp;&nbsp;正确率：
<%
			if cjs=0 then
				response.write "<font color=red>0.00%</font>"
			else
				session("oklu")=round(session("ok")*100/cjs,2)	'本次答题正确率
				response.write "<font color=red>"&session("oklu")&"%</font>"
			end if
%>		
			</td>
	</tr>
	<tr>
		<td width="70" align="center" class=tablebody1><b>奖&nbsp;&nbsp;&nbsp;金</b></td>
		<td class=tablebody1>&nbsp;现金 <font color="#FF0000">- <%=session("ajj")%> </font>&nbsp;经验 <font color="#FF0000">- <%=AnswerSetting(2)%> </font>&nbsp;魅力 <font color="#FF0000">- <%=AnswerSetting(1)%> </font></td>
		<td class=tablebody1>&nbsp;<font color="#FF0000">×&nbsp;<b>没有答对，继续努力哦！</b></font></td> 
		<td width="85" align="center" class=tablebody1><a href="z_TestReport.asp?ID=<%=session("aid")%>" title="发现该题的答案有误，我要向管理员举报" target="_blank"><font color=blue>举报此题</font></a></td>
	</tr>
	<tr>
		<td align="center" class=tablebody2><b>题&nbsp;&nbsp;&nbsp;目</b></td> 
		<td colspan="3" align="center" class=tablebody2><table border=0 width="100%" cellspacing=1><tr><td><%if cint(AnswerSetting(8))=1 then%><%=DvBCode(session("atitle"),1,2)%><%else%><%=htmlencode(session("atitle"))%><%end if%></td></tr></table></td>
	</tr>
	<tr>
		<td align="center" class=tablebody1>正确答案</td>
		<td colspan="3" class=tablebody1><%if cint(AnswerSetting(5))=1 then%><font color="#FF0000"><%=session("akey"&session("aok"))%></font><%else%>&nbsp;还要自己多学习哟！&nbsp;&nbsp; <a href="#" onclick="javascript:window.open('z_testansw.asp?aok=<%=session("aid")%>', 'newwindow', 'height=150, width=250, top=250, left=200, toolbar=no, menubar=no, scrollbars=no, resizable=no,location=no, status=no')"><font color=fuchsia>花钱买知识吧（点击此链花 <%=AnswerSetting(6)%> 元看正确答案）</font></a><%end if%></td>
	</tr> 
	</table>
	
	<tr><td class=tablebody2 align="center">
		<a href="?action=dt">继续下一题</a>　　<a href="?action=out">不玩了！</a>    
	</td></tr>
<%  
	end if 
	session("aid")=""		'题目ID  
end sub

'-------------------回答--填空题---------------
sub hdtk() 
	if session("checkpage")="" or session("aid")="" then
		call out() 
	end if
	if cint(AnswerSetting(7))=1 then	'答题是否限制时间 
		if timer()-session("ChenTimeStart")>session("atime")+5 then		'第二层判断是否回答超时，杜绝后退前进的非法答题
			response.Write "<script>alert('您的回答已经超出了本题的时间限制！')</script>"
			response.write "<script>top.location.href='?action=dt'</script>"
		end if
	end if
	
	dim cjs		'已经回答的题数
	cjs=AnswerSetting(4)-ajs	'cjs－已经回答的题数	ajs－剩余的题数

	if request("testansw")=session("akey1") then		'回答正确 
		conn.execute ("update [user] set userWealth=userWealth+"&session("ajj")&",userCP=userCP+"&AnswerSetting(1)&",userEP=userEP+"&AnswerSetting(2)&" where username='"&membername&"'")
		aconn.execute ("update [testuser] set userans=userans+1,usersinew=usersinew-"&AnswerSetting(0)&",userok=userok+1 where username='"&membername&"'")
		aconn.execute ("update [test] set tcount=tcount+1,okcount=okcount+1 where id="&session("aid")) 
%>
	<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:80%">
	<tr>
		<td align="center" class=tablebody2><b>您的属性</b></td>
		<td colspan="3" class=tablebody2>游戏币：<%=userp-AnswerSetting(0)%><font color=green>↓</font>&nbsp;&nbsp;经验：<%=myuserep+AnswerSetting(2)%><font color=red>↑</font>&nbsp;&nbsp;魅力：<%=myusercp+AnswerSetting(1)%><font color=red>↑</font>&nbsp;&nbsp;现金：<%=mymoney+session("ajj")%><font color=red>↑</font> 
		<hr size="1">出题数：<%=AnswerSetting(4)%> 题&nbsp;&nbsp;已回答：<font color=blue><%=cjs%></font> 题&nbsp;&nbsp;答对：<font color=red><%=session("ok")+1%></font> 题&nbsp;&nbsp;正确率：
<%
		if cjs=0 then
			response.write "<font color=red>0.00%</font>"
		else
			session("ok")=session("ok")+1					'本次答题正确数
			session("oklu")=round(session("ok")*100/cjs,2)	'本次答题正确率
			response.write "<font color=red>"&session("oklu")&"%</font>"
		end if
    %>		
		</td>
	</tr>
	<tr>
		<td width="70"  align="center" class=tablebody1><b>奖&nbsp;&nbsp;&nbsp;金</b></td>
		<td class=tablebody1>&nbsp;现金 <font color="#FF0000">+ <%=session("ajj")%> </font>&nbsp;经验 <font color="#FF0000">+ <%=AnswerSetting(2)%> </font>&nbsp;魅力 <font color="#FF0000">+ <%=AnswerSetting(1)%> </font></td>
		<td class=tablebody1>&nbsp;<font color=blue>√&nbsp;<b>您的回答完全正确！</b></font></td>
		<td width="85" class=tablebody1 align="center"><a href="z_TestReport.asp?ID=<%=session("aid")%>" title="发现该题的答案有误，我要向管理员举报" target="_blank"><font color=blue>举报此题</font></a></td>
	</tr>
	<tr>
		<td align="center" class=tablebody2><b>题&nbsp;&nbsp;&nbsp;目</b></td>
		<td colspan="3" class=tablebody2><%=session("atitle")%></td>
	</tr>
	<tr>
		<td align="center" class=tablebody1>正确答案</td>
		<td colspan="3" class=tablebody1>&nbsp;<%=session("akey1")%> 　</td>
	</tr>
	</table>
	
	<tr>
		<td class=tablebody2 align="center"><a href="?action=dt">继续下一题</a>　　<a href="?action=out">不玩了！</a></td>
	</tr>
<%  else		'回答错误	
		conn.execute ("update [user] set userWealth=userWealth-"&session("ajj")&",userCP=userCP-"&AnswerSetting(1)&",userEP=userEP-"&AnswerSetting(2)&" where username='"&membername&"'")
		aconn.execute ("update [testuser] set userans=userans+1,usersinew=usersinew+"&AnswerSetting(0)&" where username='"&membername&"'")
		aconn.execute ("update [test] set tcount=tcount+1 where id="&session("aid"))
%>
	<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:80%">
	<tr>
		<td align="center" class=tablebody2><b>您的属性</b></td>
		<td colspan="3" class=tablebody2>游戏币：<%=userp-AnswerSetting(0)%><font color=green>↓</font>&nbsp;&nbsp;经验：<%=myuserep-AnswerSetting(2)%><font color=blue>↓</font>&nbsp;&nbsp;魅力：<%=myusercp-AnswerSetting(1)%><font color=blue>↓</font>&nbsp;&nbsp;现金：<%=mymoney-session("ajj")%><font color=blue>↓</font>&nbsp;&nbsp;&nbsp;&nbsp;
		<hr size="1">出题数：<%=AnswerSetting(4)%> 题&nbsp;&nbsp;已回答：<font color=blue><%=cjs%></font> 题&nbsp;&nbsp;答对：<font color=red><%=session("ok")%></font> 题&nbsp;&nbsp;正确率：
<%
		if cjs=0 then
			response.write "<font color=red>0.00%</font>"
		else
			session("oklu")=round(session("ok")*100/cjs,2)	'本次答题正确率
			response.write "<font color=red>"&session("oklu")&"%</font>"
		end if
%>
		</td>
	</tr>
	<tr>
		<td width="70" class=tablebody1 align="center"><b>奖&nbsp;&nbsp;&nbsp; 金</b></td>  
		<td class=tablebody1>&nbsp;现金 <font color="#FF0000">- <%=session("ajj")%> </font>&nbsp;经验 <font color="#FF0000">- <%=AnswerSetting(2)%> </font>&nbsp;魅力 <font color="#FF0000">- <%=AnswerSetting(1)%> </font></td>
		<td class=tablebody1>&nbsp;<font color="#FF0000">×&nbsp;<b>没有答对，同志还需努力！</b></font></td> 
		<td width="85" class=tablebody1 align="center"><a href="z_TestReport.asp?ID=<%=session("aid")%>" title="发现该题的答案有误，我要向管理员举报" target="_blank"><font color=blue>举报此题</font></a></td> 
	</tr>
	<tr>
		<td align="center" class=tablebody2><b>题&nbsp;&nbsp;&nbsp; 目</b></td>
		<td colspan="3" class=tablebody2><%=session("atitle")%></td>
	</tr>
	<tr>
		<td align="center" class=tablebody1>正确答案</td>
		<td colspan="3" class=tablebody1><%if cint(AnswerSetting(5))=1 then%><font color="#FF0000"><%=session("akey1")%></font><%else%><span lang="zh-cn">&nbsp;还要自己多学习哟！&nbsp;&nbsp;		<a href="#" onclick="javascript:window.open('z_testansw.asp?aok=<%=session("aid")%>', 'newwindow', 'height=150, width=250, top=250, left=200, toolbar=no, menubar=no, scrollbars=no, resizable=no,location=no, status=no')"><font color=fuchsia>花钱买知识吧（点击此链花 <%=AnswerSetting(6)%> 元看正确答案）</font></a></span><%end if%></td>
	</tr>
	</table>
  
	<tr><td class=tablebody2 align="center">
		<a href="?action=dt">继续下一题</a>　　<a href="?action=out">不玩了！</a>
	</td></tr>
<%  
	end if
	session("aid")=""
end sub
'------------------不玩了－－退出答题----------------------
sub out()
	session("checkpage")=""
	session("ajj")=""
	session("atitle")=""
	session("akey1")=""
	session("akey1")=""
	session("akey2")=""
	session("akey3")=""
	session("akey4")=""
	session("aok")=""
	session("aid")=""
	session("testlb")=""
	session("atime")=""
	session("atx")=""
	session("acount")=""
	session("aokcount")=""
	aconn.execute ("update [testuser] set intime=now() where username='"&membername&"'")     '我来了添加
	response.write "<script>top.location.href='z_test.asp?action=TestOut'</script>"
end sub			
%>
<script language="VBScript"> 
TimeLimit=<%=session("atime")%>
function GetSec()
	TimeLimit=TimeLimit-1
	document.form1.timespent.value = TimeLimit
	if TimeLimit<= 0 then
		msgbox "你已超出本题的时间限制！",64,"开心辞典"
		top.location.replace="?action=dt"
	else
		setTimeout "GetSec()",1000 
	end if
end function
</script>
<script language="JavaScript">
<!--
	document.oncontextmenu=new Function('event.returnValue=false;'); 
	document.onselectstart=new Function('event.returnValue=false;'); 
//-->
</script>