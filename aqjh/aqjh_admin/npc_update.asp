<%@ LANGUAGE=VBScript codepage ="936" %>
<LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
userid=Request.Form("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM npc where id="&userid
rs.Open sqlstr,conn,1,2
if Request.Form("submit")="����" then rs.AddNew
if Request.Form("submit")="ɾ��" then
sqlstr="delete * from npc where id="&userid
conn.Execute sqlstr
Response.Write "�ɹ�ɾ��id="&userid&"��Npc"
else
hy=0
for i=1 to rs.Fields.Count-1
if Request.Form(i+1)="" then
	Response.Write "<script Language=Javascript>alert('��ʾ��["&rs.Fields(i).Name&"]�����ݲ���Ϊ�գ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End 
end if
if rs.Fields(i).type=11 and lcase(Request.Form(i+1))<>"true" and lcase(Request.Form(i+1))<>"false" then 
	Response.Write "<script Language=Javascript>alert('��ʾ��["&rs.Fields(i).Name&"]Ϊ�߼����磺True��False��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End 
end if
if rs.Fields(i).type=3 and (not isnumeric(Request.Form(i+1))) then 
Response.Write "<script Language=Javascript>alert('��ʾ��["&rs.Fields(i).Name&"]��ʹ���������룡');location.href = 'javascript:history.go(-1)';</script>"
Response.End 
end if
if rs.Fields(i).type=135 and (not isdate(Request.Form(i+1))) then 
Response.Write "<script Language=Javascript>alert('��ʾ��["&rs.Fields(i).Name&"]�����������Ͳ���ȷ��');location.href = 'javascript:history.go(-1)';</script>"
Response.End 
end if
if rs.fields(i).Type=202 then strtype="�ַ�"
if rs.fields(i).Type=3 then strtype="����"
if rs.fields(i).Type=135 then strtype="����"
if rs.fields(i).Type=11 then strtype="�߼�"
Response.Write "<font size=-1>"&rs.Fields(i).Name&"(<font color=blue>"&strtype&"</font>)��"&rs.Fields(i).Value&"---->"&Request.Form(i+1)&"<font color=blue>(����������ȷ!)</font></font><br><br>"
if rs.Fields(i).type=202 or rs.Fields(i).type=203 then
rs.Fields(i).Value=cstr(Request.Form(i+1))
elseif rs.Fields(i).type=3 then
rs.Fields(i).Value=clng(Request.Form(i+1))
elseif rs.Fields(i).Type=7 then
rs.Fields(i).Value=cdate(Request.Form(i+1))
end if	
next
rs.Update
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
��ϲ��npc�����޸���ɣ�
<a href="npc_list.asp">����</a>