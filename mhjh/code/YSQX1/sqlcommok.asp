<%@ LANGUAGE=VBScript%>
<%
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"

sqlstr=Request.Form("sqlstr")
fs=int(Request.form("sqlfs"))
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
Set Rs=conn.Execute(sqlstr)
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ����ִ�е�ָ��ִ����ɣ���ȷ�����أ�');</script>"
%>
<script language="vbscript">
location.href = "javascript:history.go(-1)"
</script>
<%
response.end
%>
