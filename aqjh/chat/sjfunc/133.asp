<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%
'��Ǯ����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=session("aqjh_name")
aqjh_grade=session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<20 then
	Response.Write "<script language=JavaScript>{alert('��ʾ:��Ҫ20�ȼ���');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
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
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>����Ǯ���ˡ�<font color=red>"+za(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function za(fn1,to1)
fn1=int(abs(fn1))
if to1=aqjh_name or to1="���" then
   Response.Write "<script language=JavaScript>{alert('��ʾ:���ܶ��Լ��ʹ�Ҳ�����');}</script>"
   response.end
end if
if fn1<=1000 then
   Response.Write "<script language=JavaScript>{alert('��ʾ:��Ҳ̫С���˰�,û��1000���������ò����ֵ�!');}</script>"
   response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
mygrade=rs("grade")
if rs("����")<fn1 then
	Response.Write "<script language=JavaScript>{alert('��ʾ:��������"&fn1&"�����ӣ�');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("����")="�ٸ�" and mygrade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ:���ǹٸ���,�Ͳ�Ҫ�ٵ�����!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ����ǳ�������Ҫ��ʲô��������!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select * from �û� where ����='" & to1 &"'",conn,2,2
if rs("�ȼ�")<30 or rs("����")<fn1*0.05 then
	Response.Write "<script language=JavaScript>{alert('��ʾ:�˼ҵȼ���������ô��,�ٸ�����һ�¾���������!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("����")="�ٸ�" and mygrade<10  then
	Response.Write "<script language=JavaScript>{alert('��ʾ:���ܹ����ٸ���Ա!С�ı�ץ!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ�����Գ�������ʲô��������!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
conn.execute "update �û� set ����=����-20,����=����-"&fn1&" where ����='" & aqjh_name &"'"
randomize timer
s=int(rnd*2+1)
tl=int(fn1*0.05)
if s=1 then
za="##<font color=#ff0000>�ó�"&fn1&"����ವ�һ����<font color=blue>"&to1&"</font>(������"&tl&")���˹�ȥ<img height=50 src=img/251.gif>���ҵ�"&to1&"ͷ��Ѫ��������,�򲻹�����Ǯ������~~~~�������ڵ������Σ��Լ��ĵ����½���20���еñ���ʧ�</font>"
 	conn.execute "update �û� set ����=����-"&tl&" where ����='" & to1 &"'"
else
za="##<font color=#ff0000>�ó�"&fn1&"����ವ�һ����<font color=blue>"&to1&"</font>���˹�ȥ<img height=50 src=img/251.gif>���ַ�̫��û�ҵ�,���½���20���£�������~~~~����û�뵽<font color=blue>"&to1&"</font>����һ�����������ţ�˫��ſ�أ����ð�ס��<font color=blue>"&fn1&"</font>�����ӣ��ۣ����ϵ��ڱ�����������Ц���У��ڳ������Ѷ�ʱ�����˻���Ϊʲô�������죡</font>"
	conn.execute "update �û� set ����=����+"&fn1&" where ����='" & to1 &"'"
end if
set rs=nothing	
conn.close
set conn=nothing
end function
%>
