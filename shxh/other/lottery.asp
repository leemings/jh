<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires =-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
today=date()
preday=dateadd("d",-1,today)
predaytype="#"&month(preday)&"/"&day(preday)&"/"&year(preday)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "lottery",conn
lname=rst("lname")
ldate=rst("ldate")
lmoney=rst("lmoney")
lnumber=rst("lnumber")
rst.Close 
ldatetype="#"&month(ldate)&"/"&day(ldate)&"/"&year(ldate)&"#"
pmoney=lmoney
if ldate<preday then
	ldate=preday
	ldatetype=predaytype
	conn.BeginTrans
	randomize()
	lnumber=StrReverse(mid(StrReverse(cstr(clng(rnd()*999999)))&"000000",1,6))
	rst.Open "select * from ���ϲ� where ����='"&lnumber&"' and ʱ��="&predaytype,conn,1,3
	if not rst.EOF then
		rst.PageSize=1
		lmn=clng(pmoney*0.5/rst.PageCount)
		do while not rst.EOF 
			conn.Execute "update �û� set ���=���+"&lmn&" where ����='"&rst("����")&"'"
			conn.Execute "update ���ϲ� set �����=�����+"&lmn&" where id="&rst("id")
			rst.MoveNext
			lmoney=lmoney-lmn
		loop
	end if	
	rst.Close
	rst.Open "select * from ���ϲ� where ���� like '"&mid(lnumber,1,5)&"#' or ���� like '#"&Mid(lnumber,2,5)&"' and ʱ��="&predaytype,conn,1,3
	if not rst.EOF then
		rst.PageSize=1
		lmn=clng(pmoney*0.2/rst.PageCount)
		do while not rst.EOF 
			conn.Execute "update �û� set ���=���+"&lmn&" where ����='"&rst("����")&"'"
			conn.Execute "update ���ϲ� set �����=�����+"&lmn&" where id="&rst("id")
			rst.MoveNext
			lmoney=lmoney-lmn
		loop
	end if	
	rst.Close
	rst.Open "select * from ���ϲ� where ���� like '"&mid(lnumber,1,4)&"##' or ���� like '#"&mid(lnumber,2,4)&"#' or ���� like '##"&mid(lnumber,3,4)&"' and ʱ��="&predaytype,conn,1,3
	if not rst.EOF then
		rst.PageSize=1
		lmn=1000
		do while not rst.EOF 
			conn.Execute "update �û� set ���=���+"&lmn&" where ����='"&rst("����")&"'"
			conn.Execute "update ���ϲ� set �����=�����+"&lmn&" where id="&rst("id")
			rst.MoveNext
			lmoney=lmoney-lmn
		loop
	end if	
	rst.Close
	conn.Execute "update lottery set ldate="&predaytype&",lmoney="&lmoney&",lnumber='"&lnumber&"'"
	if conn.Errors.Count=0 then
		conn.CommitTrans
	else	
		conn.RollbackTrans
		Response.Redirect "error.asp?id=104&errormsg="&conn.Errors(0).Description
	end if	
end if
rst.Open "select * from ���ϲ� where ʱ��="&ldatetype&" and �����>0 order by ����� desc",conn
do while not rst.EOF
	num=rst("����")
	if num=lnumber then 
		msg1=msg1&"<tr bgcolor=f7f7f7><td>һ�Ƚ�</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td align=right>"&rst("�����")&"</td></tr>"
	elseif mid(num,1,5)=mid(lnumber,1,5) or mid(num,2,5)=mid(lnumber,2,5) then
		msg2=msg2&"<tr bgcolor=d7d7d7><td>���Ƚ�</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td align=right>"&rst("�����")&"</td></tr>"
	else
		msg3=msg3&"<tr bgcolor=a7a7a7><td>���Ƚ�</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td align=right>"&rst("�����")&"</td></tr>"
	end if	
	rst.MoveNext 
loop
rst.Close
set rst=nothing
conn.Close 
set conn=nothing
if msg1="" then msg1="<tr bgcolor=f7f7f7><td>һ�Ƚ�</td><td colspan=2>һ�Ƚ�������,��������</td><td>�ʳ��ܽ����50%</td></tr>"
if msg2="" then msg2="<tr bgcolor=d7d7d7><td>���Ƚ�</td><td colspan=2>���Ƚ�������,��������</td><td>�ʳ��ܽ����20%</td></tr>"
if msg3="" then msg3="<tr bgcolor=a7a7a7><td>���Ƚ�</td><td colspan=2>���Ƚ�������,��������</td><td>�̶�1000Ԫ</td></tr>"
Response.Write "<head><link rel='stylesheet' href='../style.css'></head><body oncentextmenu=self.event.returnValue=false bgcolor='"&bgcolor&"' background='"&bgimage&"'><table width='100%' border=3><tr bgcolor=FFFF00><td><a href='selectnum.asp' onmouseover="&chr(34)&"window.status='����"&lname&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&lname&"</a></td><td>���ڿ���ʱ��</td><td>"&formatdatetime(ldate,1)&"</td><td>�ʳ��ܽ���</td><td>"&lmoney&"</td><td>���ڴ󽱺���</td><td>"&lnumber&"</td></tr></table><hr><table width='80%' border=3 align=center><tr bgcolor=FFFF00 align=center><td>�н�</td><td>����</td><td>����</td><td>����</td></tr>"&msg1&msg2&msg3&"</table><p align=center><a href='#' onclick='javascript:window.close();'>�ر�</a></p></body>"
%>