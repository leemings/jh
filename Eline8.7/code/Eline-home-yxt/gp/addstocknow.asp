<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻����վ�����㲻�ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
stock=Trim(Request.Form("stock"))
price=Request.Form("price")
num=Request.Form("num")
on error resume next
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../gupiao/stock.mdb")
rs.Open "select * from ��Ʊ where ��Ʊ����='"&stock&"'",conn
if not(rs.EOF or rs.BOF) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���Բ���,������Ϊ�����¹�ȡ�����ֺ���?');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.Close
set rs=nothing
conn.BeginTrans
conn.Execute "Insert Into ��Ʊ(��Ʊ����,��ͨ������,ʣ��ɷ�,���м�,��ʷ�߼�,��ʷ�ͼ�,����ɽ���,�ּ�) values('"&stock&"',"&num&","&num&","&price&","&price&","&price&","&price&","&price&")"
if conn.Errors.Count>0 then
	conn.RollbackTrans
	Response.Write "<script Language=Javascript>alert('��ʾ�����ݿ����?');location.href = 'javascript:history.go(-1)';</script>"
	Response.end 
else
	conn.CommitTrans
end if
conn.Close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ����Ʊ������ɣ�');location.href = 'javascript:history.go(-1)';</script>"
%>