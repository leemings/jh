<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����ʹ��������ս��');}</script>"
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
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="���" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>��������ս��</font><font color=" & saycolor & ">"+mpopen(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'������ս
function mpopen(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,3,3
life1=rs("����")
pai1=rs("����")
 if pai1="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('���ǹٸ����ˣ���ô�ܸ��������ɽ�Թ�������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if pai1="" or pai1="��" then
       Response.Write "<script language=JavaScript>{alert('����������,��ô����������սѽ��ȥ�������ɰɣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if life1<>"����" then
       Response.Write "<script language=JavaScript>{alert('����������,û��Ȩ���������ɲ���������ս��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
rs.open "select ����,���� FROM �û� WHERE ����='" & to1 &"'",conn
life2=rs("����")
pai2=rs("����")
 if pai2="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('�Է��ǹٸ�����,��������,��Զ��ȥ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if pai2="" or pai1="��" then
       Response.Write "<script language=JavaScript>{alert('�Է���������,��ô����������սѽ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if life2<>"����" then
       Response.Write "<script language=JavaScript>{alert('�Է���������,û��Ȩ�������Է����ɲ���������ս��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
  rs.close
       sql="SELECT q,f FROM p WHERE  a='" & pai1 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('���������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  enemy1=rs("q")
  protect1=rs("f")
    if InStr("|" & enemy1 & "|", "|" & pai2& "|")>0 then
Response.Write "<script language=JavaScript>{alert('�����Ѿ���ս��,����Ҫ����ս�˰�.');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
       sql="SELECT q,f FROM p WHERE  a='" & pai2 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('�Է���������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
 enemy2=rs("q")
  protect2=rs("f")
   if InStr("|" & enemy2 & "|", "|" & pai1& "|")>0 then
Response.Write "<script language=JavaScript>{alert('�Է������Ѿ����㿪ս��,����Ҫ����ս�˰�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       enemys1=enemy1&pai2&"|"
       enemys2=enemy2&pai1&"|"
   conn.execute"update p set q='"&enemys1&"' where a='" & pai1 & "'"
    conn.execute"update p set q='"&enemys2&"' where a='" & pai2 & "'"
mpopen=aqjh_name & "</font></b>��<b><font color=black>" & to1 & "</font></b>˵�����������µ��ӷ���ս��,������������ս,��ɱ���һͳ������ҵ." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>