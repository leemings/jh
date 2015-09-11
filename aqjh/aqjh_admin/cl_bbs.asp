<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=486"
Response.write "正在清理......<br>"
Response.Flush
Set conn=Server.CreateObject("ADODB.CONNECTION")
set rs=server.createobject("adodb.recordset") 
conn.open Application("aqjh_usermdb")
'连接论坛数据库
tdb="../bbs/database/aqjh_bbs.asp"
user_count=0
Set tconn = Server.CreateObject("ADODB.Connection")
Set jhrs=Server.CreateObject("ADODB.RecordSet")
tconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(tdb)
tconn.Open tconnstr
tsql="select * from [user]"
jhrs.open tsql,tconn,1,3
do while not jhrs.eof and not jhrs.bof
   username=jhrs("username")
   sql="select * from 用户 where 姓名='"&username&"'"
   rs.open sql,conn,1,1
   if rs.eof then
      	tsql2="delete * from [user] where username='"&username&"'"
	tconn.Execute(tsql2)
        tsql3="delete * from [forum] where username='"&username&"'"
	tconn.Execute(tsql3)
        tsql4="delete * from [favorites] where username='"&username&"'"
	tconn.Execute(tsql4)
        user_count=user_count+1
   end if
   rs.close
jhrs.movenext 
Response.Flush
loop 
says="<font color=black>【清理论坛】</font><font color=red>站长[ " & aqjh_name &" ]对论坛用户作出清理，所有在江湖中找不到资料的用户全部删除！共"&user_count&"个</font>"
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
<br>论坛用户已经清理完毕