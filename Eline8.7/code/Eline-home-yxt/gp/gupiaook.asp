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
gupiao=request.form("gupiao")
zl=request.form("zl")
sy=request.form("sy")
fx=request.form("fx")
gj=request.form("gj")
dj=request.form("dj")
cjj=request.form("cjj")
xj=request.form("xj")
if gupiao="" or zl="" or sy="" or fx="" or gj="" or dj="" or cjj="" or xj="" then
Response.write "����δ��д����!����������"
Response.end 
end if
if request.form("submit1")="ɾ��" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../gupiao/stock.mdb")
conn.execute("delete from ��Ʊ where ��Ʊ����='"&gupiao&"'")
conn.execute("delete from �ֹ� where ����='"&gupiao&"'")
conn.close
set conn=nothing
Response.write "���ݿ������ɣ����Ѿ��ɹ�ɾ���˹�Ʊ"
Response.end 
end if
if request.form("submit")="����" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../stock/stock.mdb")
conn.execute("update ��Ʊ set ��ͨ������='"&zl&"',ʣ��ɷ�='"&sy&"',���м�='"&fx&"',��ʷ�߼�='"&gj&"',��ʷ�ͼ�='"&dj&"',����ɽ���='"&cjj&"',�ּ�='"&xj&"' where ��Ʊ����='"&gupiao&"'")
conn.close
set conn=nothing
Response.write "���ݿ������ɣ����Ѿ��ɹ������˹�Ʊ"
Response.end 
end if
%>