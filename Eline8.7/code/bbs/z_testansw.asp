<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_testconn.asp"-->
<%
	if (not founduser) or membername="" then
		response.Write "<br><br><p align=center><font color=red>您怎么会跑到这里来了，请不要从外部提交本页</font></p><br><br>"
		response.end
	end if
	dim aok,ars
	aok=request("aok")
	set ars=server.createobject("adodb.recordset")
	sql="select key1,key2,key3,key4,ok,tx from [test] where id="&aok
	ars.open sql,aconn,1,1
	if ars.eof and ars.bof then
		response.Write"<br><br><p align=center><font color=red>题目不存在啊，请不要从外部提交本页</font></p><br><br>"
		response.end
	else
		if ars(5)=0 then 
			select case ars(4)
				case "1"
					aok=ars(0)
				case "2"
					aok=ars(1)
				case "3"
					aok=ars(2)
				case "4"
					aok=ars(3)
			end select
		else
			aok=ars(0)
		end if
	end if
	ars.close
	dim sellmoney
	sellmoney=split(aconn.execute("select AnswerSetting from [config]")(0),"||")(6)
	conn.execute ("update [user] set userWealth=userWealth-"&sellmoney&" where username='"&membername&"'")
	set aconn=nothing
	set conn=nothing
%>
<br><p align="center"><font color=red>正确答案：<%=aok%></font></p>
<br><p align="center"><a href="#" onclick=window.close()>关闭</a></p>