<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'���˿�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "../chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))

'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
says=Ucase(says)
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=dpkask(fnn1,aqjh_name,towho)
says="<font color=green><b>�����˿ˡ�</b><font color=" & saycolor & ">"&says&"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function dpkask(fn1,from1,to1)
nickname=Session("aqjh_name")
grade=Session("aqjh_grade")
chatroomsn=session("nowinroom")

if to1="���" or to1=from1 then call ErrALT("�㲻����ѡ���һ�������Ϊ���֣�")

'------------------------�¸�ʽ------------------------
ARR_FN=Split(fn1,"|")
dim ub_Err,xia_class,xiamoney,xia_fir,xia_sql
ub_Err="�ڲ�������������ʽ���£�\n /���˿� ��ע��Ŀ|����[�򣺽�ң�����] \n\n��ʾ���ۣ�\n /���˿� 1000|����\n /���˿� 1000|���\n /���˿� 1000|����\n /���˿� 1000|���� "

if ubound(ARR_FN)<>1 Then call ErrALT(ub_Err)
xia_class=ARR_FN(1)
xiamoney=ARR_FN(0)

select case xia_class
	case "���"
		xia_fir="1"
		xia_sql="���"
	case "����"
		xia_fir="2"
		xia_sql="����"
	case "����"
		xia_fir="3"
		xia_sql="�ݶ�����"
	case else
		call ErrALT(ub_Err)
end select

If not isnumeric(xiamoney) Then call ErrALT(ub_Err)

xiamoney=abs(int(xiamoney))

if (xia_fir="1")and(xiamoney<3 or xiamoney>10) then Call ErrALT("����Ѻ��3" & xia_class & "�����10����ң�")
if (xia_fir="2")and(xiamoney<1000 or xiamoney>50000000) then Call ErrALT("����Ѻ��1000" & xia_class & "�����5000��������")
if (xia_fir="3")and(xiamoney<100 or xiamoney>3000) then Call ErrALT("����Ѻ��100" & xia_class & "�����5000���㣡")

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
'------------------------�¸�ʽ------------------------
sql="select " & xia_sql & " from �û� where ����='" &from1&"'"
set rs=conn.execute(sql)
yin1=rs(0)
rs.close
if xiamoney> yin1 then
	set rs=nothing
	conn.close
	set conn=nothing
	Call ErrALT("������Ϻ���û����ô��" & xia_class)
end if
sql="select " & xia_sql & " from �û� where ����='" &to1&"'"
set rs=conn.execute(sql)
yin2=rs(0)
rs.close
if xiamoney>yin2 then
	set rs=nothing
	conn.close
	set conn=nothing
	Call ErrALT("�Է����Ϻ���û����ô��" & xia_class)
end if
'------------------------�¸�ʽ------------------------
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="DPK.ASP"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�����ķ�����֧��jet4.0����ʹ�����ص����ӷ�������߳�������
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
sql="select * from dpk where ufrom='"& nickname & "' or uto='"& nickname &"'"
Set Rs=conn.Execute(sql)
if not (rs.eof and rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script Language='Javascript'>alert('�������ƾ�û�н��������ܿ���');parent.f3.location.href='dpk-xp.asp';</script>"
	response.end
end if
rs.close
sql="select * from dpk where ufrom='"& to1 & "' or uto='"& to1 &"'"
Set Rs=conn.Execute(sql)
if not (rs.eof and rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script Language='Javascript'>alert('"&to1&"�����ƾ�û�н��������ܿ���');parent.f3.location.href='dpk-xp.asp';</script>"
	response.end
end if
rs.close
'------------------------�¸�ʽ------------------------
dpkask="<b><font color=green>["&from1&"]</font></b>��<b><font color=black>["&to1&"]</font></b>˵�����˿��𣿶�עΪ<b><font color=red>"&xiamoney & xia_class & "</font></b>��<img src='f2/dpk/1.GIF'><input type=button value='Ը��' onclick=""javascript:window.open('f2/dpkok.asp?id="&myid&"&from1="&from1&"&to1="&to1&"&yn=1','a','width=380,height=210');this.disabled=1"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><img src='f2/dpk/2.GIF'><input type=button value='�ܾ�' onclick=""javascript:window.open('f2/dpkok.asp?id="&myid&"&from1="&from1 & "&to1="&to1&"&yn=0','a','width=380,height=210');this.disabled=1"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223"">"
sql="insert into dpk(ufrom,uto,duz) values ('"& nickname & "$��ע','"& to1 & "$��ע', " & xia_fir & xiamoney & ")"
Set Rs=conn.Execute(sql)
conn.close
set rs=nothing
set conn=nothing
end function
%>
