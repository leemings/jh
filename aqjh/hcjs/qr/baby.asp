<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �Ա�,����,hw from �û� where ����='"&aqjh_name&"'",conn
if rs("�Ա�")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��һ�������˽���������ʲô����ȥ��');window.close();}</script>"
	Response.End
end if
if rs("hw")<>"" then
	rs.close
	conn.execute "update �û� set ����='����' where ����='"&aqjh_name&"'"
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('Ҫ�븻������·���������Ӷ����������Ѿ��к����ˣ�����������');window.close();}</script>"
	Response.End
end if
if rs("����")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��û���о��������ӣ���');window.close();}</script>"
	Response.End
end if
yun=rs("����")
rs.close
hdata=split(yun,"|")
baby=hdata(0)
babysj=hdata(1)
babytl=clng(hdata(2))
babyzt=hdata(3)
erase hdata
sysj=DateDiff("d",babysj,now())
if sysj<20 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���в���20�죬δ������ʱ�䡣');window.close();}</script>"
	Response.End
end if
if sysj>21 then
	conn.execute "update �û� set ����='����' where ����='"&aqjh_name&"'"
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㻳��ʱ�䳬��21�죬�����������ڸ����ˣ�');window.close();}</script>"
	Response.End
end if
newbl="����"
randomize timer
r=int(rnd*5)+1
if sysj=20 or (sysj=21 and r<>5 and r<>6) then
	n=int(rnd*3)+1
	if n=1 or n=3 then
		xhxb="��"
		h1="����С��"
		t="boy.gif"
	else
		xhxb="Ů"
		h1="Ư������"
		t="girl.gif"
	end if
	xhtl=babytl+1000
	xhnl=int(xhtl*3/4)
	xhz="С��|"&xhxb&"|"&xhtl&"|"&xhnl&"|200|200|"&now()&"|��|0|"&now()&"|0"
	hh="<img src=images/"&t&"><BR><BR><img src='img/004.gif'><br><marquee width=100% behavior=alternate scrollamount=5><font color=red size=+1>ϲϲ</font></marquee><BR>��ϲ�㣬��˳������һ��"& h1 &"������������"& xhtl &"��������"&xhnl&"��������200��������200����ȥ���ߺ��ӵİְ֣�����һ���������������������ǿ��ֵĽᾧ��"
	mess="��ϲ<font color=blue>"& aqjh_name &"</font>����һ��<font color=red><b>"& h1 &"</b></font><BR><marquee width=100% behavior=alternate scrollamount=5><font color=red size=+1>ϲϲ</font></marquee>"
else
	hh="�����������ڵ�20���ʱ��δ�ܼ�ʱ��ҽԺ��������һ���ʱ�䡣��Ȼ����ҽ����Ŭ��������ĺ��ӻ��������ˡ��Բ�����ڰ�˳�㡣"
	xhz=""
	mess="����"& aqjh_name &"����ʱ�䳬��20�죬������С��ʱ�����Ѳ���С�������ˣ�Ĭ������"
end if
conn.execute "update �û� set ����='"&newbl&"',hw='"&xhz&"' where ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000><b>��ҽԺ��Ϣ��</b></font>"&mess			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session(nowinroom) & ");<"&"/script>"
addmsg saystr
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
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
  <p> <br>
  </p>
  <p><%=hh%> </p>
  <p>
    <input type=button value=�رմ��� onClick='window.close()' name="button">
  </p>
</div>
</body></html>