<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("Ba_jxqy_username")
if un="" then Response.Redirect "../error.asp?id=016"
mg=Request.QueryString("mg")
if instr(mg,"'")<>0 then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 体力,内力,特效 from 物品 where 名称='"&mg&"' and 所有者='"&un&"' and 数量>0",conn
if rst.EOF or rst.BOF then
	usecurative="你并没有"&mg&"可供使用"
else
	adhp=0-rst("体力")
	admp=rst("内力")
	adespecial=mid(rst("特效"),2,2)
	if isnull(adespecial) then adespecial="无"
	conn.execute "update 用户 set 体力=体力-"&adhp&",内力=内力+"&admp&" where 姓名='"&un&"'"
	conn.execute "update 物品 set 数量=数量-1 where 名称='"&mg&"' and 所有者='"&un&"'"
	msg="<font color=FF0000>【用药】</font>##对%%使用了"&mg&",使之体力增加"&abs(adhp)&",内力增加"&admp
	usecurative="你对"&st&"成功使用了"&mg
	aberrantname=Application("Ba_jxqy_aberrantname")
	if instr(aberrantname,";"&st&";")<>0 and adespecial<>"无" then
		aberrantlist=Application("Ba_jxqy_aberrantlist")
		aberrantlistubd=ubound(aberrantlist)
		for i=1 to aberrantlistubd step 4
			if aberrantlist(i)=st and aberrantlist(i+2)=adespecial then aberrantlist(i+3)=dateadd("s",adhp,aberrantlist(i+3))
		next
		Application.Lock
		Application("Ba_jxqy_aberrantlist")=aberrantlist
		Application.UnLock
		erase aberrantlist
		aberranttxt="，并使之"&adespecial&"程度减轻"&abs(adhp)
	end if
	usecurative="你成功使用了"&mg&aberranttxt
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
	newtalkarr(593)=1
	newtalkarr(594)=un
	newtalkarr(595)=un
	newtalkarr(596)=""
	newtalkarr(597)="#660099"
	newtalkarr(598)="#660099"
	newtalkarr(599)="<font color=FF0000>【用药】<\/font>##对%%使用了"&mg&",使之体力增加"&adhp&",内力增加"&admp&aberranttxt&"<font class=timsty>（"&time()&"）<\/font>"
	newtalkarr(600)=chatroomsn
	Application.Lock
	Application("Ba_jxqy_talkpoint")=talkpoint+1
	Application("Ba_jxqy_talkarr")=newtalkarr
	Application.UnLock
	erase newtalkarr
	erase talkarr
end if
rst.Close
set rst=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>使用药品</title>
<link rel=stylesheet href='../chatroom/css.css'>
<script language=javascript>
setTimeout("location.href='seecurative.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<div align=center>
<font color=0000ff size=5>使用药品</font>
<hr>
三秒钟自动返回<br>
<input type=button value='返回' onclick="javascript:location.href='seecurative.asp'">
</div>
<%=usecurative%>
</body>
</html>