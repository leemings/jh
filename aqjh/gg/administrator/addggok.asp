<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if aqjh_grade<>10 or aqjh_name<>application("aqjh_user") then Response.Redirect "../../error.asp?id=439"
gg=trim(request.form("gg"))
bt=trim(request.form("bt"))
lb=trim(request.form("lb"))
if bt="" then
	response.write "<script language=JavaScript>{alert('��������д������⣡��');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
if gg="" then
	gg="������"
end if
lbdata=split(lb,"|")
lbz=clng(lbdata(0))
lbm=lbdata(1)
gg=replace(gg,chr(13),"<br>")
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
conn.execute "insert into ����(����,ʱ��,����,�������,�������,�鿴����) values ('"&bt&"',now(),'"&gg&"',"&lbz&",'"&lbmc&"',0)"
set rs=nothing
conn.close
set conn=nothing
response.write "<script language=JavaScript>{alert('���������ϣ���');location.href = 'ggadmin.asp';}</script>"
response.end
%>
