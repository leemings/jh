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
co=Session("Ba_jxqy_usercorp")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 等级,消耗精力,特效 from 攻击 where 姓名='"&un&"' and 招式='"&mg&"'",conn
if rst.EOF or rst.BOF then
	rst.Close
	rst.Open "select * from 招式 where 门派='"&co&"' and 招式='"&mg&"'",conn
	if rst.EOF or rst.BOF then
		msg="<font color=FF0000>【修习】</font>##想修习"&mg&"，而苦无高人指点，因而失败。"
		learn="本门没有此样招式可供修习！"
	else
		energy=rst("消耗精力")
		proviso=rst("修习条件")
		basemp=rst("消耗内力")
		baseat=rst("基本攻击")
		especial=rst("特效")
		atdeclaration=rst("攻击说明")
		rst.Close
		rst.Open "select * from 用户 where 姓名='"&un&"' and 精力>="&energy&" and "&proviso,conn
		if rst.EOF or rst.BOF then
			provisotxt=replace(proviso,"and","并且")
			provisotxt=replace(provisotxt,"or","或者")
			provisotxt=replace(provisotxt,">=","不低于")
			provisotxt=replace(provisotxt,"<=","不高于")
			provisotxt=replace(provisotxt,">","大于")
			provisotxt=replace(provisotxt,"<","小于")
			provisotxt=replace(provisotxt,"=","为")
			msg="<font color=FF0000>【修习】</font>##修习"&mg&"时，虽有高人指点，苦于自己资质有限。因而失败。"
			learn="修习时失败，需精力"&energy&"，条件满足："&provisotxt
		else
			conn.Execute "insert into 攻击(姓名,招式,等级,消耗精力,消耗内力,基本攻击,特效,攻击说明) values('"&un&"','"&mg&"',1,"&energy&","&basemp&","&baseat&",'"&especial&"','"&atdeclaration&"')"
			conn.execute "update 用户 set 精力=精力-"&energy&" where 姓名='"&un&"'"
			msg="<font color=FF0000>【修习】</font>##修习"&mg&"第1级成功。耗用精力<bgsound src='../mid/xiulian.wav' loop=1>"&energy
			learn="恭喜您修习"&mg&"入门，以后你可以自行修行了"
		end if	
	end if
elseif rst("等级")<100 then
	energy=rst("消耗精力")
	agrade=rst("等级")
	rst.Close
	rst.Open "select * from 用户 where 姓名='"&un&"' and 精力>="&energy*agrade,conn
	if rst.EOF or rst.BOF then
		msg="<font color=FF0000>【修习】</font>##修习"&mg&"第"&agrade&"级时心有旁骛失败，无法专心入定，因而失败"
		learn="修习时失败，需精力"&energy
	else
		conn.Execute "update 用户 set 精力=精力-"&energy*agrade&" where 姓名='"&un&"'"
		conn.Execute "update 攻击 set 等级=等级+1,基本攻击=基本攻击*1.1 where 姓名='"&un&"' and 招式='"&mg&"'"
		msg="<font color=FF0000>【修习】</font>##修习"&mg&"第"&agrade&"级成功。耗用精力<bgsound src='../mid/xiulian.wav' loop=1>"&energy*agrade
		learn="恭喜您修习"&mg&"至"&agrade+1&"，以后你可以自行修行了"
	end if
else	msg="<font color=FF0000>【修习】</font>##你已经修习"&mg&"至最高了无法再行修练。<bgsound src='../mid/xiulian.wav' loop=1>"
	learn="我们确信您在"&mg&"是高手中的高手了，再修习也不会有什么效果了！"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg=replace(msg,"\","\\")
msg=replace(msg,"/","\/")
msg=replace(msg,chr(34),"\"&chr(34))
msg=replace(msg,chr(39),"\"&chr(39))
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
newtalkarr(599)=msg&"<font class=timsty>（"&time()&"）<\/font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("Ba_jxqy_talkpoint")=talkpoint+1
Application("Ba_jxqy_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<html>
<head>
<title>修习武功</title>
<link rel=stylesheet href='css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<div align=center>
<font color=0000ff size=5>修习武功</font>
<hr>
<input type=button value='返回' onclick="javascript:location.href='learn.asp'">
</div>
<%=learn%>
</body>
</html>