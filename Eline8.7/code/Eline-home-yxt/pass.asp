<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
cpass=Request.Form("cpass")
if cpass=Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('��ʾ����վ�������㲻���Ըģ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="SELECT * FROM �û� where ����='"&cpass&"'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('��Ǹ����Ҫ���ҵ��������Ҳ�������鿴�Ƿ���ȷ��');history.back();</script>"
else
sql="update �û� set ����='e10adc3949ba59abbe56e057f20f883e' where ����='"&cpass&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','������¼','�������룺["&cpass&"]��Ϊ123456��')"
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=javascript>alert('�û���["&cpass&"]�������޸ĳ�123456�ɹ���');history.back();</script>"
end if
%>