<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���ѡ�wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����ѡ�<font color=" & saycolor & ">"+haoyou(towho,mid(says,i+1))+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'����
function haoyou(to1,fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select �������� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
pengyou=trim(rs("��������"))
fn1=trim(fn1)
select case fn1
case "����"
if instr(LCase(pengyou),LCase(to1))<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��["& to1 &"]�Ѿ�����ĺ����ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
pengyou=pengyou& "|" &to1& "|"
conn.execute "update �û� set ��������='"& pengyou &"' where ����='" & sjjh_name &"'"
Response.Write "<script language=JavaScript>{alert('��ʾ��["& to1 &"]�����Ѿ�����!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End

case "ɾ��"
if instr(pengyou,trim(to1))=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��["& to1 &"]��������ĺ��ѣ����޷�ɾ����');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
towho="|" &trim(to1)& "|"
pengyou=Replace(pengyou,towho,"")
if pengyou="" then
	pengyou=" "
end if
conn.execute "update �û� set ��������='"& pengyou &"' where ����='" & sjjh_name &"'"
Response.Write "<script language=JavaScript>{alert('��ʾ��["& to1 &"]�����Ѿ�ɾ��!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End

case "�鿴"
if pengyou="" then
	Response.Write "<script language=JavaScript>{alert('��ʾ��["& sjjh_name &"]�㲢û�к��ѣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
haoyou="##��������:"&pengyou
end select
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
