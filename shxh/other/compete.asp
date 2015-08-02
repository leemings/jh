<%
Response.Expires=-1
username=session("Ba_jxqy_username")
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
	elseif horsetmp>9999 then
		horsetmp=9999
	end if
	horse(i)=horsetmp
	horsepoint(i)=600
	horselist=horselist&horsepoint(i)&","
	horsecalc=horsecalc+horse(i)
next
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
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
conn.close
set conn=nothing
erase horsepoint
msg=msg&"<head><title></title><STYLE>"
for i=0 to 11
	msg=msg&"#Layer"&i&" {position: absolute;HEIGHT: 60px;WIDTH: 60px; Z-INDEX:3;left:600px;top="&i*30&"px}"
next
msg=msg&"</style><script language=javascript>var pos=0;var i=0;var posxlist='"&horselist&"';var posx=posxlist.split(/,/gi);function start(){for(i=0;i<12;i++){if(posx[pos+i]>=0){eval('document.all["&chr(34)&"Layer'+i+'"&chr(34)&"].style.left=posx[pos+i]');}else{win();return true;}}pos=pos+12;setTimeout('start()',10);}function win(){parent.betfrm.location.replace('win.asp?calc="&horsecalc&"&win="&horsewin&"&bet="&horse(horsewin)&"');}</script></head><body oncontextmenu=self.event.returnValue=false bgcolor='"&bgcolor&"' background='"&bgimage&"'"
if horsecalc<>0 then msg=msg&" onload=start()"
msg=msg&">"
for i=0 to 11
msg=msg&"<DIV id=Layer"&i&"><dd><IMG  src='../images/horse.gif' ></dd></DIV>"
next
msg=msg&"<div style='position: absolute;z-index:2'><table  border=1 bgcolor="&bgcolor&" width='100%' border=1>"
for i=0 to 11 
   msg=msg&"<tr><td width=32 height=18 align=right>"&i&"号</td><td width=632 height=18 bgcolor="&rgb((11-i)*20,(11+i)*7,i*20)&">&nbsp;</td><td> 赌金<input type=text readonly maxlength=4 size=4 value='"&horse(i)&"'></td></tr>"
next
msg=msg&"</table></div></body>"
erase horse
Response.Write msg
%>
