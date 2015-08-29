<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="z_jifen.inc"-->
<!--#include file="z_plus_check.asp"-->
<%
	stats="保存积分"
	call nav()
	call head_var(2,0,"","")

dim msgtext,savestats

if membername="" or not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛积分的权限，请先登录或者同管理员联系。"
	call dvbbs_error()	
else
	call main()
end if
call activeonline()
call footer()

'下面是主程序
sub main()
 dim username
 username=membername
 dim stat,online,onlineT
 dim Srs
 set rs=server.createobject("adodb.recordset")
 sql="select * from [User] where username='"&username&"'"
 rs.open sql,conn,1,1
 if err.number<>0 then 
 	savestats="失败"
    msgtext="<br><br><br>数据库操作失败："&err.description&"<BR><BR><BR>"
	call message()
    exit sub
 end if

 if rs.eof and rs.bof then
 	savestats="失败"
	msgtext="<br><br><br>用户不存在! <BR><BR><BR><BR><BR>"
	call message()
  exit sub
 else
	dim srss,Q,W,C,T
	set Srs=conn.execute("select stats,startime from online where username='"&username&"'")
	if Srs.eof and Srs.bof then        '用户在线数据未写入online表中或因超时已从online表中删除  
		msgtext="<br><img src=face/face4.gif border=0> 发生意外<br>"
	  msgtext=msgtext&"<br><br>============= <font color=red>意外原因</font> =============<br>"
	  msgtext=msgtext&"<br><li>刚登录，用户的在线数据未写入在线数据表中"
	  msgtext=msgtext&"<br><li>因 "&Forum_Setting(8)&" 分钟没有活动而被服务器从表中删除 <br><br>"
	  call message()
	  exit sub  
	end if 
	
	online=datediff("n",Srs(1),Now()) '提取在线时间  online=在线时间（分）
	onlineT=online + 5
	T=jf_StartTime            '在线时间（分）后开始计算

	select case SuanfaValue
		case 1
			if jf_WealthRate=0 then
				Q=0
			else
				Q=Int(online / jf_WealthRate) 
			end if
			if jf_EPRate=0 then
				W=0
			else
				W=Int(online / jf_EPRate) 
			end if
			if jf_CPRate=0 then
				C=0
			else
				C=Int(online / jf_CPRate)
			end if
		case 2
			if jf_WealthRate=0 then
				Q=0
			else
				Q=Int(onlineT / jf_WealthRate) 
			end if
			if jf_EPRate=0 then
				W=0
			else
				W=Int(onlineT / jf_EPRate) 
			end if
			if jf_CPRate=0 then
				C=0
			else
				C=Int(onlineT / jf_CPRate)
			end if
		case else
			if jf_WealthRate=0 then
				Q=0
			else
				Q=iif(Int(onlineT / jf_WealthRate)>0,Int(onlineT / jf_WealthRate),1) 
			end if
			if jf_EPRate=0 then
				W=0
			else
				W=iif(Int(onlineT / jf_EPRate)>0,Int(onlineT / jf_EPRate),1) 
			end if
			if jf_CPRate=0 then
				C=0
			else
				C=iif(Int(onlineT / jf_CPRate)>0,Int(onlineT / jf_CPRate),1) 
			end if
	end select
	
	dim rsl,sqll
	sqll="select * from [User] where  username='"&username&"'"
		set rsl=server.createobject("adodb.recordset")
	 rsl.open sqll,conn,1,3
	
	'userCP     魅力值
	'userEP     经验值
	'userWealth 财产值
	
	if online>=T then
		sqll="update [user] set userWealth=userWealth+"&Q&",userEP=userEP+"&W&",userCP=userCP+"&C&" where username='"&username&"'"
		conn.execute(sqll)
		rsl.close
		set rsl=nothing
	'保存当前时间
		set Srss=conn.execute("update [online] set startime=Now() where username='"&username&"'") 
		Srs.close
		set Srs=nothing
		rs.close
		set rs=nothing
		savestats="成功"
		msgtext="<br><img src=face/face14.gif border=0> <font color=blue>"&username&"</font>,恭喜你"
		msgtext=msgtext&"<br><br>============= <font color=red>积分情况</font> =============<br><br>"
		msgtext=msgtext&username&" 你在线 "&online&" 分钟 <br><br>"
		if jf_WealthRate>0 then
			msgtext=msgtext&username&" 增加了 "&Q&" 个金币 <br>"
		end if
		if jf_EPRate>0 then
			msgtext=msgtext&username&" 增加了 "&w&" 点经验 <br>"
		end if
		if jf_CPRate>0 then
			msgtext=msgtext&username&" 增加了 "&c&" 点魅力 <br>"
		end if
		msgtext=msgtext&"<br><br>============= <font color=red>积分规则</font> =============<br>"
		msgtext=msgtext&"<br><li>在线 "&T&" 分钟以上才可以保存积分&nbsp;&nbsp;"
		msgtext=msgtext&"<br><li>保存积分的时间间隔为 "&T&" 分钟&nbsp;&nbsp;&nbsp;&nbsp;"
		msgtext=msgtext&"<br><li>系统不提供自动保存积分功能&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		msgtext=msgtext&"<br><li>积分保存后在线时间重零开始计算&nbsp;&nbsp;"		
		msgtext=msgtext&"<br><br>============= <font color=red>计算方法</font> =============<br>"
		if jf_WealthRate>0 then
			msgtext=msgtext&"<br><li>财产值 = 在线时间 / "&jf_WealthRate
		end if
		if jf_EPRate>0 then
			msgtext=msgtext&"<br><li>经验值 = 在线时间 / "&jf_EPRate
		end if
		if jf_CPRate>0 then
			msgtext=msgtext&"<br><li>魅力值 = 在线时间 / "&jf_CPRate
		end if
		call message()
	else
		savestats="时间未到"
		msgtext="<br><img src=face/face23.gif border=0> 这么快就想来拿工资啦？时间还没有到啊！<br><br>"
		msgtext=msgtext&username&" 你才上线 "&online&" 分钟 <br>"
		msgtext=msgtext&"<br><br>============= <font color=red>积分规则</font> =============<br>"
		msgtext=msgtext&"<br><li>在线 "&T&" 分钟以上才可以保存积分&nbsp;&nbsp;"
		msgtext=msgtext&"<br><li>保存积分的时间间隔为 "&T&" 分钟&nbsp;&nbsp;&nbsp;&nbsp;"
		msgtext=msgtext&"<br><li>系统不提供自动保存积分功能&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		msgtext=msgtext&"<br><li>积分保存后在线时间重零开始计算&nbsp;&nbsp;"		
		msgtext=msgtext&"<br><br>============= <font color=red>计算方法</font> =============<br>"
		if jf_WealthRate>0 then
			msgtext=msgtext&"<br><li>财产值 = 在线时间 / "&jf_WealthRate
		end if
		if jf_EPRate>0 then
			msgtext=msgtext&"<br><li>经验值 = 在线时间 / "&jf_EPRate
		end if
		if jf_CPRate>0 then
			msgtext=msgtext&"<br><li>魅力值 = 在线时间 / "&jf_CPRate
		end if
		call message()
		exit sub
	end if 
end if
end sub

sub message()
%>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
<tr><th>保存积分</th></tr>
<tr><td class=tablebody2>

     <table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:350">
        <tr>
          <th height=20 >保存积分 － <%=savestats%></th>
        </tr>
        <tr>
		<td nowrap align="center" class=tablebody1>
			<%=msgtext%>
          </td>
        </tr>
        <tr>
          <td nowrap align="center" class=tablebody2>
			<a href="#" onClick="window.open('z_jifenhelp.htm','积分帮助','scrollbars=yes,menubar=no,resizable=yes,top=25,left=100,width=400,height=500')"><font color=red>>> 积分帮助 <<</font></a>
          </td>
       </tr>
      </table>
    </td></tr>
<tr><th align="center"><a href=javascript:history.go(-1)> <font color=#ffffff><< 返回上一页</font> </a></th></tr>
</table>	  
<%
end sub

function iif(expr,truepart,falsepart)
	'函数功能：如果表达式 expr 为真，则返回 truepart 的值；否则返回 falsepart 的值
	if expr then 
		iif=truepart
	else
		iif=falsepart
	end if	
end function
%>
