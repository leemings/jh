<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
username=Request.Form("search")
cz=Request.form("sjcz")
name=request.querystring("username")
if name<>"" then
username=name
cz="����"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if cz="ID" then
username=int(username)
sqlstr="SELECT * FROM �û� where "&cz&"="&username
rs.Open sqlstr,conn
else
sqlstr="SELECT * FROM �û� where "&cz&"='"&username&"'"
rs.Open sqlstr,conn
end if
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('��Ǹ����Ҫ���ҵ��������Ҳ�������鿴�Ƿ���ȷ��');history.back();</script>"
else
if rs("grade")=10 and aqjh_name<>Application("aqjh_user") then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������޸�վ�����ϣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','���ң�["&cz&"]="&username&"�ļ�¼��')"
Response.Write "<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666><LINK href=css/css.css type=text/css rel=stylesheet>"
Response.Write "<form method=post action=updateuser.asp><table  border=0 cellspacing=1 cellpadding=4 bgcolor=#f2f2ea height='20' align='center'><tr><td colspan='6'><div align='center'>����ID:"&rs("id")&"<input type=hidden readonly size=4 name=id value="&rs("id")&">(ע�⣺��������������Ӧ�����ݱ�����ͬ��������������пյ����ݴ���!)</div></td></tr>"
Response.Write "<tr>"
for i=1 to rs.Fields.Count-25
strtype=rs.fields(i).Type
strname=rs.Fields(i).Name
if strname="����" then seename=rs.Fields(i).value
if rs.fields(i).Type=202 then strtype="�ַ�"
if rs.fields(i).Type=3 then strtype="����"
if rs.fields(i).Type=7 then strtype="����"
if rs.fields(i).Type=11 then strtype="�߼�"
if strname="grade" then strname="����ȼ�"
if strname="mvalue" then strname="�»���"
if strname="allvalue" then strname="�ܻ���"
if strname="times" then strname="��¼����"
if strname="regip" then strname="ע��ip"
if strname="lastip" then strname="���ip"
if strname="regtime" then strname="ע��ʱ��"
if strname="lasttime" then strname="���ʱ��"
'if strname<>"cw" and strname<>"w1" and strname<>"w2"  and strname<>"w3"  and strname<>"w4"  and strname<>"w5"  and strname<>"w6"    and strname<>"w7"   and strname<>"w8" and strname<>"z1"   and strname<>"z2"   and strname<>"z3"   and strname<>"z4"   and strname<>"z5"   and strname<>"z6"  then
	Response.Write "<td bgcolor='#f2f2ea'><div align='right'>"&strname&"(<font color=blue>"&strtype&"</font>)��<input type=text size='10' name="&rs.Fields(i).Name&" value='"&rs.Fields(i).value&"' class=form style='BORDER-BOTTOM:#B8AF86 1px solid'></font></div></td>"
	if i/3=int(i/3) then Response.Write "</tr><tr>"
'end if
	next
Response.Write "</tr>"
Response.Write "<tr><td colspan='6'><div align='center'><input type=submit value=���� name=submit> <input type=submit value=���� name=submit> <input type=submit value=ɾ�� name=submit> <input type=reset value=����> <a href='../chat/TOWUPIN.asp?toname="&seename&"' target='_blank' >��ѯ������Ʒ</a></div></td></tr>"
Response.Write "</table></form>"
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
