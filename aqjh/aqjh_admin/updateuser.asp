<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
userid=Request.Form("id")
Response.Write "<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM �û� where id="&userid
rs.Open sqlstr,conn,1,2
if Request.Form("submit")="����" then rs.AddNew
if Request.Form("submit")="ɾ��" then
sqlstr="delete * from �û� where id="&userid
conn.Execute sqlstr
Response.Write "�ɹ�ɾ��id="&userid&"���û�"
else
hy=0
dim nc(16)
nc(1)="w1"
nc(2)="w2"
nc(3)="w3"
nc(4)="w4"
nc(5)="w5"
nc(6)="w6"
nc(7)="w7"
nc(8)="w8"
nc(9)="w9"
nc(10)="w10"
nc(11)="z1"
nc(12)="z2"
nc(13)="z3"
nc(14)="z4"
nc(15)="z5"
nc(16)="z6"
for i=1 to rs.Fields.Count-25
myform=rs.Fields(i).Name
'���
for j=1 to 16
 if myform=nc(j) then
	chkok="yes"
	exit for
 else
	chkok="no"
 end if
next
if Request.Form(i+1)="" and chkok="no" then
	Response.Write "<script Language=Javascript>alert('��ʾ��["&rs.Fields(i).Name&"]�����ݲ���Ϊ�գ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End 
end if
'���
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

if rs.Fields(i).Name="��Ա�ȼ�" then hy=clng(Request.Form(i+1))
if rs.Fields(i).Name="��Ա����" and hy>0 then
	if DateDiff("d",date(),cdate(Request.Form(i+1)))<10 then
		Response.Write "<script Language=Javascript>alert('��ʾ��["&rs.Fields(i).Name&"]��Ա���ڲ���ȷ������\n�û�Ա�ȼ�ʱ����Ա����Ӧ����10�죡');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
end if
if rs.Fields(i).Name="����" then mp=cstr(Request.Form(i+1))
if rs.Fields(i).Name="���" then sf=cstr(Request.Form(i+1))
if lcase(rs.Fields(i).Name)="grade" then dj=clng(Request.Form(i+1))
if dj>10 then
	Response.Write "<script Language=Javascript>alert('��ʾ������Ա��ߵȼ�Ϊ10����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End 
end if
if dj<>0 then
	if mp="�ٸ�" and dj<6 then
		Response.Write "<script Language=Javascript>alert('��ʾ��������Ϊ�ٸ�ʱ��[����]Ӧ����6��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if mp<>"�ٸ�" and dj>=6 then
		Response.Write "<script Language=Javascript>alert('��ʾ�������ɲ�Ϊ�ٸ�ʱ��[����]ӦС��6��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if sf="����" and mp<>"�ٸ�" and dj<>5 then
		Response.Write "<script Language=Javascript>alert('��ʾ��[���]���ŵĹ���ȼ�Ϊ5����');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if sf="����" and mp<>"�ٸ�" and dj<>4 then
		Response.Write "<script Language=Javascript>alert('��ʾ��[���]���ϵĹ���ȼ�Ϊ4����');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if sf="����" and mp<>"�ٸ�" and dj<>3 then
		Response.Write "<script Language=Javascript>alert('��ʾ��[���]�����Ĺ���ȼ�Ϊ3����');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
	if sf="����" and mp<>"�ٸ�" and dj<>2 then
		Response.Write "<script Language=Javascript>alert('��ʾ��[���]�����Ĺ���ȼ�Ϊ2����');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
end if
if rs.Fields(i).Name="��̸���" and Request.Form(i+1)<>"/" and Request.Form(i+1)<>"����" and Request.Form(i+1)<>"վ��" then
		Response.Write "<script Language=Javascript>alert('��ʾ��[��̸���]ֻ��Ϊ��/ ���� վ����/��ʾ�ɷ���������');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
end if
if rs.Fields(i).Name="����" then yin=clng(Request.Form(i+1))
if rs.Fields(i).Name="���" then 
	yin=yin+clng(Request.Form(i+1))
	if yin>2000000000 then
		Response.Write "<script Language=Javascript>alert('��ʾ��[���������]�ܺϲ����Գ���20�ڣ���');location.href = 'javascript:history.go(-1)';</script>"
		Response.End 
	end if
end if
Response.Write rs.Fields(i).Name&"(<font color=blue>"&strtype&"</font>)��"&rs.Fields(i).Value&"---->"&Request.Form(i+1)&"<font color=blue>(����������ȷ!)</font></font><br><br>"
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
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','["&Request.Form("submit")&"]ID="&userid&"�ļ�¼��')"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<LINK href="css/css.css" type=text/css rel=stylesheet>
��ϲ���û������޸���ɣ�
<a href="fine.asp">����</a>