<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then 
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻����վ�����㲻�ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="SELECT * FROM b WHERE ��Ʒ��='" & name & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
if clng(Request.Form("wupinll"))<0 then Response.Redirect "../error.asp?id=471"
if clng(Request.Form("wupintl"))<0 then Response.Redirect "../error.asp?id=471"
name=Request.Form("wupinname")
a=clng(Request.Form("wupinll"))
b=clng(Request.Form("wupintl"))
wupinll=(a+b)*500
set rs=server.createobject("adodb.recordset")
sql="select * from b where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("a")=name
rs("b")=request.form("lx")
rs("f")=a
rs("g")=b
rs("h")=wupinll
rs.update
rs.close
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=700"
%>
<%else%>
<script language=vbscript>
MsgBox "�ֿ��Ѿ��������Ʒ�ˣ�������ӵ��İ�"
location.href = "javascript:history.back()"
</script>
<%end if%>
