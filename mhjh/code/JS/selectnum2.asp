<%
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
dim num(4)
money=0
for i=0 to 4
num(i)=Request.Form("num"&i&"0")&Request.Form("num"&i&"1")&Request.Form("num"&i&"2")&Request.Form("num"&i&"3")&Request.Form("num"&i&"4")&Request.Form("num"&i&"5")
if num(i)<>"" then
money=money+100000
if len(num(i))<>6 then Response.Redirect "../error.asp?id=220"
if not isnumeric(num(i)) then Response.Redirect "../error.asp?id=220"
msg=msg&"<tr><td>"&num(i)&"</td><td align=right>100</td></tr>"
end if
next
bgcolor=Application("yx8_mhjh_backgroundcolor")
today=date()
preday=dateadd("d",-1,today)
todaytype="#"&month(today)&"/"&day(today)&"/"&year(today)&"#"
predaytype="#"&month(preday)&"/"&day(preday)&"/"&year(preday)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&username&"' and ����>="&money,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
conn.BeginTrans
rst.Open "select * from ���ϲ� where ʱ��<"&predaytype,conn
for i=0 to 4
if num(i)<>"" then
if rst.EOF or rst.BOF then
conn.Execute "insert into ���ϲ�(����,����,ʱ��,�����) values('"&username&"','"&num(i)&"',"&todaytype&",0)"
else
conn.Execute "update ���ϲ� set ����='"&username&"',����='"&num(i)&"',ʱ��="&todaytype&",�����=0 where id="&rst("id")
rst.MoveNext
end if
end if
next
rst.Close
set rst=nothing
conn.Execute "update �û� set ����=����-"&money&" where ����='"&username&"'"
conn.Execute "update lottery set lmoney=lmoney+"&money
conn.CommitTrans
conn.Close
set conn=nothing
Response.Write "<head><link rel='stylesheet' href='../style.css'></head><body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif'><table width='80%' border=3  bgcolor=ffffff align=center><tr bgcolor=ffffff align=center><td>��Ч����</td><td>��ע</td></tr>"&msg&"<tr><td>����ע</td><td align=right>"&money&"</td></tr></table><p align=center><a href='#' onclick='javascript:history.back();'>����</a></p><p align=center>���Ѿ������˸�����ȯ�������쿪���ɣ�лл�����ˣ�</p></body>"
%>
