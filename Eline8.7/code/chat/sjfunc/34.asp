<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�����wWw.51eline.com��
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
says="<font color=green>��������ѡ�</font><font color=" & saycolor & ">"+clspy(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'���
function clspy(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
pengyou=trim(rs("��������"))
fn1=trim(fn1)
if instr(pengyou,trim(fn1))=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["& fn1 &"]��������ĺ��ѣ����޷�ɾ����');}</script>"
	Response.End
end if
towho="|" &trim(fn1)& "|"
pengyou=Replace(pengyou,towho,"")
if pengyou="" then
	pengyou=" "
end if
conn.execute "update �û� set ��������='"& pengyou &"' where ����='" & sjjh_name &"'"
Response.Write "<script language=JavaScript>{alert('["& fn1 &"]�Ѿ��Ӻ����б������!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
clspy="##���ĺ���:"&fn1&"�Ѿ����б������������"
end function
%>
