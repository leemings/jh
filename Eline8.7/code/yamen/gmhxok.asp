<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("sjjh_jm")=session("sjjh_jm")+1
if session("sjjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
if Weekday(date())=1 and Hour(time())>=20 and Hour(time())<21 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ϊ����ʱ�䣬���ɸ�����');window.close();}</script>"
	Response.End 
end if
myid=Request.form("id")
name=lcase(trim(request.form("name")))
pass=request.form("pass")
name1=lcase(trim(request.form("name1")))
if not isnumeric(myid) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&myid&"]�������ID��ʹ�����֣�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
for i=1 to len(name1)
usernamechr=mid(name1,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=120"
next
if name1="��" or name="��" or name="δ��" or name1="δ��" then Response.Redirect "../error.asp?id=130"
if left(name1,3)="%20" OR InStr(name1,"=")<>0 or InStr(name1,"`")<>0 or InStr(name1,"'")<>0 or InStr(name1," ")<>0 or InStr(name1,"��")<>0 or InStr(name1,"'")<>0 or InStr(name1,chr(34))<>0 or InStr(name1,"\")<>0 or InStr(name1,",")<>0 or InStr(name1,"<")<>0 or InStr(name1,">")<>0 then Response.Redirect "../error.asp?id=120"
if chuser(name1) then Response.Redirect "../error.asp?id=120"
if instr(name,"or")<>0 or instr(name,"'")<>0 or instr(name,"|")<>0 or instr(name," ")<>0 then Response.Redirect "../error.asp?id=120"
if instr(name1,"or")<>0 or instr(name1,"'")<>0 or instr(name1,"|")<>0 or instr(name1," ")<>0 then Response.Redirect "../error.asp?id=120"
if Instr(name1,"��ϯ")>0 or Instr(name1,"����")>0 or Instr(name1,"��������")>0 or Instr(name1,"��������")>0 or Instr(name1,Application("sjjh_automanname"))>0 or Instr(name1,"һ����")>0 or Instr(name1,"վ��")>0 or Instr(name1,"��Ȼ")>0 or Instr(name1,"����")>0 or Instr(name1,"ʱ��")>0 or Instr(name1,"����")>0 or Instr(name1,"�й�")>0 or Instr(name1,"��")>0 or Instr(name1,"��")>0 or Instr(name1,"��")>0 or Instr(name1,"���")>0 or Instr(name1,"��")>0 then Response.Redirect "../error.asp?id=130"
if name="" or name1="" or pass="" then Response.Redirect "../error.asp?id=56"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(name)&" ")<>0 then Response.Redirect "../error.asp?id=61"
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'У���û�
rs.open "SELECT * FROM �û� WHERE ����='" & name1 & "'",conn
If not rs.bof and not rs.eof  Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����û����Ѿ����ڣ�');history.go(-1)</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM �û� WHERE id="&myid&" and ����='" & name & "' and ����='" & pass & "'",conn
If Rs.Bof OR Rs.Eof Then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��ԭ�û����Ѿ����ڻ����벻��ȷ��');history.go(-1)</script>"
	Response.End
end if
if rs("�ȼ�")<60 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����������û��ðɣ�����60�����ܸ�����');window.parent.close();</script>"
	Response.End
end if
if rs("����")<500000000 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��ûǮҲ���������׬�����������ɣ�');window.parent.close();</script>"
	Response.End
end if
if rs("���")<200 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��û��һ����������׬���������ɣ�');window.parent.close();</script>"
	Response.End
end if
if rs("���")="����" or rs("����")="�ٸ�" or rs("grade")>=6 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǹٸ����˻��������������ţ���ֹ������');window.parent.close();</script>"
	Response.End
end if
sex=rs("�Ա�")
bl=rs("��ż")
qren=rs("����")
qren2=rs("��������")
qren3=rs("��������")
conn.Execute "update �û� set ����='" & name1 & "',����=����-500000000,���=���-200,��Ա�ȼ�=��Ա�ȼ�-1 where ����='" & name & "'"
conn.Execute "update �û� set ʦ��='" & name1 & "' where ʦ��='" & name & "'"
conn.Execute "update �û� set ������='" & name1 & "' where ������='" & name & "'"
if qren<>"��" and qren<>"" then
	conn.Execute "update �û� set ����='" & name1 & "' where ����='" & qren & "'"
end if
if qren2<>"��" and qren2<>"" then
	conn.Execute "update �û� set ��������='" & name1 & "' where ����='" & qren2 & "'"
end if
if qren3<>"��" and qren3<>"" then
	conn.Execute "update �û� set ��������='" & name1 & "' where ����='" & qren3 & "'"
end if
if bl<>"��" and bl<>"" then
	conn.Execute "update �û� set ��ż='" & name1 & "' where ����='" & bl & "'"
	if sex="��" then
		conn.execute "update t set b='"& name1 &"' where b='" & name & "'"
	else
		conn.execute "update t set c='"& name1 &"' where c='" & name & "'"
	end if
end if
conn.execute "update k set a='"& name1 &"' where a='" & name & "'"
conn.execute "update j set a='"& name1 &"' where a='" & name & "'"
conn.execute "update n set b='"& name1 &"' where b='" & name & "'"
conn.execute "update vh set b='"& name1 &"' where b='" & name & "'"
'��¼
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& name &"','"& name1 &"','����','�������֣�')"
rs.close
set rs=nothing
conn.close
set conn=nothing

'��¼
Set conn2=Server.CreateObject("ADODB.CONNECTION")
conn2.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../gupiao/eline_stock.asa")
conn2.Execute "update �� set �ʺ�='" & name1 & "' where �ʺ�='" & name & "'"
conn2.Execute "update �ͻ� set �ʺ�='" & name1 & "' where �ʺ�='" & name & "'"
conn2.Execute "update ��Ʊ set ��Ӫ��='" & name1 & "' where ��Ӫ��='" & name & "'"
conn2.close
set conn2=nothing
Response.Write "<script Language=Javascript>alert('��ϲ[" & name1 & "]�����ָ�����ɣ������ھͿ�����������½�ˣ�');window.parent.close();</script>"
response.end
%>