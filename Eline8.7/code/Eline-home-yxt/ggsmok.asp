<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
cz=trim(request.querystring("cz"))
hysm=Trim(Request.Form("hysm"))
hydz=Trim(Request.Form("hydz"))
if cz="����" and Instr(hysm,";")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��ʽ���£�\n���1;���2;���3;');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if cz="����" and Instr(hysm,";")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��ʽ���£�\n����1;����2;����3;');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if cz="����" and Instr(hydz,";")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��ʽ���£�\n��1|��1;��2|��2;��3|��3;');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if cz="����" and Instr(hydz,"|")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��ʽ���£�\n��1|��1;��2|��2;��3|��3;');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sqlstr="SELECT * FROM sm where a='"&cz&"'"
rs.Open sqlstr,conn,1,2
rs("c")=hysm
if cz="���" then
	rs("b")=clng(hydz)
else
	rs("d")=hydz
end if
rs.Update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ�������޸ĳɹ�!');location.href = 'ggsm.asp';</script>"
%>
