<%
Response.Expires=-1
lastip=Request.ServerVariables("LOCAL_ADDR")
templastip=StrReverse(left(lastip&"9fhb6utime,./",15))
templenlastip=len(lastip)
mmcheck=""
for j=1 to 15
mmcheck=mmcheck+chr(asc(mid(templastip,j,1))-templenlastip+int(j*1.1))
next
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
dim horse(11)
horsecalc=0
horselist=""
dim horsepoint()
for i=0 to 11
redim preserve horsepoint(i)
horsetmp=Trim(Request.Form("horse"&i))
if horsetmp="" then
horsetmp=0
elseif not isnumeric(horsetmp) then
horsetmp=0
elseif horsetmp<0 then
horsetmp=0
elseif horsetmp>999999 then
horsetmp=999999
end if
horse(i)=horsetmp
horsepoint(i)=600
horselist=horselist&horsepoint(i)&","
horsecalc=horsecalc+horse(i)
next
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from 用户 where 姓名='"&username&"' and 银两>="&horsecalc
rst.open sqlstr,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
set rst=nothing
randomize()
do while true
redim preserve horsepoint(i+1)
horsepoint(i)=horsepoint(i-12)-clng(rnd()*5)
horselist=horselist&horsepoint(i)&","
if horsepoint(i)<0 then
horsewin=i mod 12
exit do
end if
i=i+1
loop
conn.Execute "update 用户 set 银两=银两+"&horse(horsewin)*10-horsecalc&" where 姓名='"&username&"'"
	kl="<font color=0000FF>"&username&"</font>在娱乐中心赌马,本次战果为<font color=0000FF>"&horse(horsewin)*10-horsecalc&"</font>两白银！"
conn.close
set conn=nothing
says="<font color=red><b>【娱乐赌博】</b></font>"&kl	
dim newtalkarr(600) 
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
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
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=name
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
erase horsepoint
msg=msg&"<head><title></title><STYLE>"
for i=0 to 11
msg=msg&"#Layer"&i&" {position: absolute;HEIGHT: 60px;WIDTH: 60px; Z-INDEX:4;left:600px;top="&i*30&"px}"
next
msg=msg&"</style><script language=javascript>var pos=0;var i=0;var posxlist='"&horselist&"';var posx=posxlist.split(/,/gi);function start(){for(i=0;i<12;i++){if(posx[pos+i]>=0){eval('document.all["&chr(34)&"Layer'+i+'"&chr(34)&"].style.left=posx[pos+i]');}else{win();return true;}}pos=pos+12;setTimeout('start()',10);}function win(){parent.betfrm.location.replace('win.asp?calc="&horsecalc&"&win="&horsewin&"&bet="&horse(horsewin)&"');}</script></head><body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif'"
if horsecalc<>0 then msg=msg&" onload=start()"
msg=msg&">"
for i=0 to 11
msg=msg&"<DIV id=Layer"&i&"><dd><IMG  src='../image/horse.gif' ></dd></DIV>"
next
msg=msg&"<div style='position: absolute;z-index:1'><table  border=1  width='100%' border=1>"
for i=0 to 11
msg=msg&"<tr><td width=32 height=18 align=right>"&i&"号</td><td width=632 height=18 bgcolor="&rgb((11-i)*20,(11+i)*7,i*20)&">&nbsp;</td><td> 赌金<input type=text readonly maxlength=4 size=4 value='"&horse(i)&"'></td></tr>"
next
msg=msg&"</table></div></body>"
erase horse
Response.Write msg
%>
<head>
<title></title>
<base target="_self">
<bgsound src="../mid/horse.MID" loop="-1">
</head>
