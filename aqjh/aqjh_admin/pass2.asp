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
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
cpass=Request.Form("cpass")
if aqjh_name<>Application("aqjh_user") then 
	Response.Write "<script Language=Javascript>alert('��ʾ��ֻ����վ�������޸ģ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="SELECT * FROM �û� where ����='"&cpass&"'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('��Ǹ����Ҫ���ҵ��������Ҳ�������鿴�Ƿ���ȷ��');history.back();</script>"
else
sql="update �û� set �ڶ�����='e10adc3949ba59abbe56e057f20f883e' where ����='"&cpass&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','���ĵڶ����룺["&cpass&"]��Ϊ123456��')"
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=javascript>alert('�û���["&cpass&"]�ĵڶ������޸ĳ�123456�ɹ���');history.back();</script>"
end if
%>