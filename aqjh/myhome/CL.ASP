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
Set conn2=Server.CreateObject("ADODB.CONNECTION")
Set rs2=Server.CreateObject("ADODB.RecordSet")
connstr2=Application("aqjh_usermdb")
conn2.open connstr2
%>
<LINK href=../aqjh_admin/css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%
sql="SELECT * FROM house"
Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
  	huzhu=rs("h_拥有者")
	banlu=rs("h_拥有者2")
	f_id=rs("id")
	sql2="select * from 用户 where h_拥有者='" & huzhu & "' and h_拥有者2='" & banlu & "'"
	Set Rs2=conn2.Execute(sql2)
	if rs2.eof then
                if huzhu<>"无" then
		sql2="update house set h_拥有者='无',h_拥有者2='无' where h_id="&f_id
		conn.Execute(sql2)
                Response.write "×"
   else
              Response.write "√"
  end if   	
else
        Response.write "√"
	end if
Rs2.close
rs.movenext 
Response.Flush
loop 
conn.close
says="<font color=black>【清理房屋】</font><font color=red>站长[ " & aqjh_name &" ]对快乐江湖房屋进行了清理，所有无主房屋都被清理完毕！</font>"			'聊天数据
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
<br>所有无主房屋清理完毕
<br>（改名,被杀或再婚导致的无主房屋全部清理并提供给大家购买）
<br>×为找不到主人的无主房屋