
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
stock=Request.form("stock")
num=Request.Form("num")
if instr(stock,"'")<>0 then 
	Response.Write "<script Language=Javascript>alert('��ʾ����������');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
if not isnumeric(num) then 
	Response.Write "<script Language=Javascript>alert('��ʾ����������');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
num=abs(int(clng(num)))
if num<1 then num=1
msg="<head><link rel='stylesheet' href='../../style.css'><script language=javascript>setTimeout('history.back();',3000);</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 bgcolor='#339966'><h3>�깺ԭʼ��</h3><hr>"
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
set conn1=server.CreateObject("adodb.connection")
set rs1=server.CreateObject("adodb.recordset")
conn1.open Application("yx8_mhjh_connstr")
rs1.Open "select ���� from �û� where ����='"&username&"'",conn1,2,2
if rs1.EOF or rs1.BOF then 
	set rs=nothing
	conn.close
	set conn=nothing
	rs1.close
	set rs1=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����ݿ��ѯ����');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
umoney=rs1("����")
rs1.Close
rs.Open "select * from ��Ʊ where ��Ʊ����='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then 
	set rs1=nothing
	conn1.close
	set conn1=nothing
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���Բ���û�д��ֹ�Ʊ');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
uprice=rs("���м�")
unum=umoney\uprice
uallownum=rs("ʣ��ɷ�")
if num>uallownum then
	set rs1=nothing
	conn1.close
	set conn1=nothing
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��ʣ���Ʊֻ��:"&uallownum&"�ɣ�');location.href = 'javascript:history.go(-1)'</script>"
	Response.end 
end if
if uallownum<unum then uallownum=unum
	if num>unum then 
		set rs1=nothing
		conn1.close
		set conn1=nothing
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ���Բ������Ǯ���󲻹�Ӵ��');location.href = 'javascript:history.go(-1)'</script>"
		Response.end 
	end if
if num>uallownum then num=uallownum
umoney=num*uprice
rs.Close
rs.Open "select * from �ֹ� where �ֹ���='"&username&"' and ����='"&stock&"'",conn,2,2
if rs.EOF or rs.BOF then
	conn.Execute "insert into �ֹ�(�ֹ���,ʱ��,����,��Ȩ,����,���,����) values('"&username&"',now(),'"&stock&"',"&num&",0,"&uprice&",0)"
else
	conn.Execute "update �ֹ� set ��Ȩ=��Ȩ+"&num&" where �ֹ���='"&username&"' and ����='"&stock&"'"
end if
	conn.Execute "update ��Ʊ set ʣ��ɷ�=ʣ��ɷ�-"&num&" where ��Ʊ����='"&stock&"'"
	conn1.Execute "update �û� set ����=����-"&umoney&" where ����='"&username&"'"
rs.Close
set rs=nothing
if conn.Errors.Count>0 then
msg=msg&"<p align=center>����<font color=FF0000>ʧ��</font>,�����Ӻ��Զ�����<br><a href='javascript:history.back();'>����</a></p><script language=javascript>parent.stockfrm.location.reload();</script></body>"
conn.RollbackTrans
else
msg=msg&"<p align=center>����<font color=0000FF>���</font>,�����Ӻ��Զ�����<br><a href='javascript:history.back();'>����</a></p><script language=javascript>parent.stockfrm.location.reload();</script></body>"
end if 
conn.Close
set conn=nothing
Response.Write msg
%>