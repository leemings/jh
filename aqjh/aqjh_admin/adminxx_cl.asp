<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Response.write "正在清理......<br>"
Response.Flush
Set conn=Server.CreateObject("ADODB.CONNECTION")
set rs=server.createobject("adodb.recordset")
conn.open Application("aqjh_usermdb")
sql="Select * From n order by b"
rs.open sql,conn,1,1
do while not rs.eof and not rs.bof 
	xxid=rs("id")
	xxb=rs("b")
        set rs2=server.createobject("adodb.recordset")
	sql2="select * from 用户 where 姓名='"&xxb&"'"
        rs2.open sql2,conn,1,1
	if rs2.eof then
		sql2="delete * from n where id="&xxid 
		conn.Execute(sql2)
        Response.write "×"
	else
        Response.write "√"
	end if
Rs2.close 
rs.movenext 
Response.Flush
loop 
conn.close
says="<font color=black>【清理武功】</font><font color=red>站长[ " & aqjh_name &" ]对快乐江湖轩辕进行了清理，所有无主的轩辕都已经清理完毕！</font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
call addmsg(saystr)
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<LINK href="css/css.css" type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<br>所有无主的轩辕清理完毕