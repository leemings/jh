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
Response.write "��������......<br>"
Response.Flush
Set conn=Server.CreateObject("ADODB.CONNECTION")
set rs=server.createobject("adodb.recordset") 
conn.open Application("aqjh_usermdb")
'������̳���ݿ�
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
   sql="select * from �û� where ����='"&username&"'"
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
says="<font color=black>��������̳��</font><font color=red>վ��[ " & aqjh_name &" ]����̳�û��������������ڽ������Ҳ������ϵ��û�ȫ��ɾ������"&user_count&"��</font>"
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
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
<br>��̳�û��Ѿ��������