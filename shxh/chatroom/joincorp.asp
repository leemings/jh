<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("Ba_jxqy_username")
co=session("Ba_jxqy_usercorp")
if un="" then Response.Redirect "../error.asp?id=016"
mg=Request.QueryString("mg")
if instr(mg,"'") then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 门派 where 门派='"&mg&"'",conn
if rst.EOF or rst.BOF then
	joincorp="您无法想加入"&mg&"，因为武林中并没有这个门派。"
else
	cond=rst("加入条件")
	rst.Close
	rst.Open "select * from 用户 where 姓名='"&un&"' and "&cond,conn
	if rst.EOF or rst.BOF then
		joincorp="您的条件不够,"&mg&"拒绝接收这样的弟子。"
	else
		if co="无" then
			conn.Execute "update 用户 set 门派='"&mg&"',身份='无' where 姓名='"&un&"'"
			onlinelist=Application("Ba_jxqy_onlinelist")
			onlinelistubd=ubound(onlinelist)
			for i=1 to onlinelistubd step 8
				if onlinelist(i)=un then
					onlinelist(i+2)=mg
					exit for
				end if
			next
			Application.Lock
			Application("Ba_jxqy_onlinelist")=onlinelist
			Application.UnLock
			erase onlinelist
			session("Ba_jxqy_usercorp")=mg
			joincorp="欢迎您加入"&mg
			talkarr=Application("Ba_jxqy_talkarr")
			talkpoint=clng(Application("Ba_jxqy_talkpoint"))
			dim newtalkarr(600)
			j=1
			for i=11 to 600 step 10
				newtalkarr(j)=talkarr(i)
				newtalkarr(j+1)=talkarr(i+1)
				newtalkarr(j+2)=talkarr(i+2)
				newtalkarr(j+3)=talkarr(i+3)
				newtalkarr(j+4)=talkarr(i+4)
				newtalkarr(j+5)=talkarr(i+5)
				newtalkarr(j+6)=talkarr(i+6)
				newtalkarr(j+7)=talkarr(i+7)
				newtalkarr(j+8)=talkarr(i+8)
				newtalkarr(j+9)=talkarr(i+9)
				j=j+10
			next
			newtalkarr(591)=talkpoint+1
			newtalkarr(592)=2
			newtalkarr(593)=0
			newtalkarr(594)=un
			newtalkarr(595)="大家"
			newtalkarr(596)=""
			newtalkarr(597)="#660099"
			newtalkarr(598)="#660099"
			newtalkarr(599)="<font color=FF0000>【加入】</font>##成功的加入"&mg&"<font class=timsty>（"&time()&"）<\/font>"
			newtalkarr(600)=chatroomsn
			Application.Lock
			Application("Ba_jxqy_talkpoint")=talkpoint+1
			Application("Ba_jxqy_talkarr")=newtalkarr
			Application.UnLock
			erase newtalkarr
			erase talkarr
		elseif co=mg then
				joincorp="您已经是"&co&"的弟子了，不用再加入"&mg
		else
			joincorp="可叹呀可悲，##身为"&co&"弟子，居然想重投"&mg
		end if	
	end if
end if
rst.Close 
set rst=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>使用药品</title>
<link rel=stylesheet href='css.css'>
<script language=javascript>
setTimeout("location.href='corp.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<div align=center>
<font color=0000ff size=5>使用药品</font>
<hr>
三秒钟自动返回<br>
<input type=button value='返回' onclick="javascript:location.href='corp.asp'" id=button1 name=button1>
</div>
<%=joincorp%>
</body>
</html>