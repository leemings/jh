<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=439"
t1=trim(Request("t1"))
t2=trim(Request("t2"))
username=Request("name")
sl=request("sl")
if t1="" or username="" or sl="" then
   Response.Write "<script language=javascript>alert('�벻Ҫ�����ֵ��');history.back();</script>"
   response.end
end if
if not isnumeric(sl) then
   Response.Write "<script language=javascript>alert('������ʹ�����֣�');history.back();</script>"
   response.end
end if
if t1="��Ա��" and sl>10000 then
   Response.Write "<script language=javascript>alert('��Ҳ̫���˰ɣ����ó���1��');history.back();</script>"
   response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM �û� where ����='"&username&"'"
rs.Open sqlstr,conn
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('��Ǹ����Ҫ�������û������ڣ���鿴�Ƿ���ȷ��');history.back();</script>"
else
 id=rs("id")
 if t2="����" then
  zt=rs(t1)+sl
 elseif t2="����" then
  zt=rs(t1)-sl
 else
  Response.Write "<script language=javascript>alert('δ֪����');history.back();</script>"
 end if
 zt=int(zt)
 if zt>100000000 then
   Response.Write "<script language=javascript>alert('���棺Ҳ̫���˰ɣ���������1���ˣ�');history.back();</script>"
   response.end
 elseif zt<0 then
   Response.Write "<script language=javascript>alert('���棺����ʹ�û�״̬Ϊ����');history.back();</script>"
   response.end
 end if
 conn.execute "update �û� set "&t1&"="&zt&" where id="&id
 conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','�޸ģ�"&username&"��["&t1&"]��¼��')"
Response.Write "<script language=javascript>alert('��ʾ���޸����ݳɹ���');window.location.href='fine.asp';</script>"
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
