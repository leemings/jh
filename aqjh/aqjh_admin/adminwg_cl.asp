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
Response.write "��������......<br>"
Response.Flush
Set conn=Server.CreateObject("ADODB.CONNECTION")
set rs=server.createobject("adodb.recordset")
conn.open Application("aqjh_usermdb")
sql="Select * From y order by b"
rs.open sql,conn,1,1
do while not rs.eof and not rs.bof 
	wgid=rs("id")
	wgb=rs("b")
        if wgb<>"����" and wgb<>"����" then
        set rs2=server.createobject("adodb.recordset")
	sql2="select * from p where a='"&wgb&"'"
        rs2.open sql2,conn,1,1
	if rs2.eof then
		sql2="delete * from y where id="&wgid 
		conn.Execute(sql2)
        Response.write "��"
	else
        Response.write "��"
	end if
        Rs2.close
        end if
rs.movenext 
Response.Flush
loop 
conn.close
says="<font color=black>�������书��</font><font color=red>վ��[ " & aqjh_name &" ]�Կ��ֽ����书�������������������Ѿ���ɾ�����书���Ѿ�������ϣ�</font>"			'��������
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
<br>���������Ѿ���ɾ�����书�������