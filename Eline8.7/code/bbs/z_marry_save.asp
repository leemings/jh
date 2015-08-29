<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<%
dim content

toname=request.form("sBride")
content=request.form("sLoveWord")
guige=request.form("iStandard")
theyear=request.form("wed_b_y")
them=request.form("wed_b_m")
thed=request.form("wed_b_d")
theh=request.form("wed_b_h")
lasttime=request.form("sLastTime")
fumoney=request.form("cPayMethod")
total=guige*100+lasttime*100+400

call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有查看本社区教堂的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	call dvbbs_error()
else
	call marry_nav()
	call main()
end if
call activeonline()
call footer()

sub main()
	dim tousermoney,yourmoney
	set rs=server.createobject("adodb.recordset")
	'rs.open"select * from [user] where username='"&membername&"'",conn,3,1
	'yourmoney=rs("userwealth")
	'rs.close
	'dim tousermoney
	'yourmoney=conn.execute("select userwealth from [user] where username='"&membername&"'")(0)
	'tousermoney=conn.execute("select userwealth from [user] where username='"&toname&"'")(0)
	'if (fumoney=1 and total>yourmoney) or (fumoney=2 and ) then 
	'	Errmsg=Errmsg+"<br>"+"<li>您的现金不足！"
	'	call dvbbs_error()
	'else
		rs.open"select * from [qiuhun] where (UserName='"&membername&"' or tousername='"&membername&"') and dlg=true",connj,3,1
		if rs.eof then
			rs.close
			Errmsg=Errmsg+"<br>"+"<li>你还没有向任何人求婚或是没人答应你的求婚！"
			call dvbbs_error()
			response.end
		elseif rs("jiehun") =true then
			rs.close
			Errmsg=Errmsg+"<br>"+"<li>你们都已经登记结婚了，怎么还来登记啊？"
			call dvbbs_error()
			response.end
		else
			if rs("username")=membername then
				ok=true
				name2=rs("tousername")
			elseif rs("tousername")=membername then
				ok=false
				name2=rs("username")
			end if
			rs.close
			if toname<>name2 then
				Errmsg=Errmsg+"<br>"+"<li>你是不是找错人了？"
				call dvbbs_error()
				response.end
			else
				yourmoney=conn.execute("select userwealth from [user] where username='"&membername&"'")(0)
				tousermoney=conn.execute("select userwealth from [user] where username='"&toname&"'")(0)
				if fumoney=1 then
					if yourmoney<total then
						ErrMsg=ErrMsg+"<br>"+"<li>您的现金不足，请选择其它付款方式！"
						call dvbbs_error()
						response.end
					else
						conn.execute"update [user] set userwealth=userwealth-"&total&",wife='"&name2&"' where username='"&membername&"' " 
						conn.execute"update [user] set wife='"&membername&"' where username='"&name2&"' " 
					end if
				elseif fumoney=2 then
					if tousermoney<total then 
						ErrMsg=ErrMsg+"<br>"+"<li>您爱人的现金不足，请选择其它付款方式！"
						call dvbbs_error()
						response.end
					else
						conn.execute"update [user] set userwealth=userwealth-"&total&",wife='"&membername&"' where username='"&name2&"' " 
						conn.execute"update [user] set wife='"&name2&"' where username='"&membername&"' " 
					end if
				elseif fumoney=3 then
					total=total/2
					if tousermoney<total or yourmoney<total then
						ErrMsg=ErrMsg+"<br>"+"<li>你们有一方现金不足总金额的一半，请选择其它付款方式！"
						call dvbbs_error()
						response.end
					else
						conn.execute"update [user] set userwealth=userwealth-"&total&",wife='"&name2&"' where username='"&membername&"' " 
						conn.execute"update [user] set userwealth=userwealth-"&total&",wife='"&membername&"' where username='"&name2&"' " 
					end if
				end if
				set rs1=server.createobject("adodb.recordset")
				if ok then
					rs1.open"select * from [jie] where username='"&membername&"' and thename='"&name2&"'",connj,3,3
				else
					rs1.open"select * from [jie] where username='"&name2&"' and thename='"&membername&"'",connj,3,3
				end if
				rs1("long")=lasttime
				rs1("content")=content
				'rs1("year")=theyear&"-"&them&"-"&thed
				rs1("year")=now()
				rs1("type")=guige
				rs1("jiehun")=true
				rs1.update
				rs1.close
				connj.execute"update [qiuhun] set jiehun=true where username='"&membername&"' or tousername='"&membername&"'" 
				%>
				<script language="vbscript">
					msgbox"祝贺你呵呵，结婚了，请你去邀请你的朋友，记得要准时哦",0,"<%=Forum_info(0)%>"
					location.href = "z_marry_index.asp"
				</script>
				<%
				set rs1 = Server.CreateObject("ADODB.recordset")
				rs1.open"Select * From [message]",conn,3,3
				rs1.addnew
				rs1("sender")="婚姻中心"
				rs1("incept")=name2
				rs1("title")="恭喜"&membername&"和你结成夫妻!"
				rs1("content")="恭喜你们于"&now()&"结成夫妻<br>请到婚姻中心去看看吧"
				rs1("sendtime")=now()
				rs1("issend")=1
				rs1.update
				rs1.addnew
				rs1("sender")="婚姻事务所"
				rs1("incept")=membername
				rs1("title")="恭喜"&name2&"和你结成夫妻!"
				rs1("content")="恭喜你们于"&now()&"结成夫妻<br>请到婚姻中心去看看吧"
				rs1("sendtime")=now()
				rs1("issend")=1
				rs1.update
				rs1.close
			end if
		end if
	'end if
end sub %>