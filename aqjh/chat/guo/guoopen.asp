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
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Թ���֮�����ƣ�');}</script>"
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
says="<font color=green>���������ơ�</font><font color=" & saycolor & ">"+guoopen(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'������ս
function mpopen(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ְλ,����,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,3,3
zhiwei1=rs("ְλ")
guo1=rs("����")
pai1=rs("����")
 if pai1="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('���ǹٸ����ˣ���ô�ܲ������֮��ķ������������Լ���ְ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if guo1="" or guo1="��" then
       Response.Write "<script language=JavaScript>{alert('��������û�м����κι���,��ô�����սѽ��ȥ����ɣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if zhiwei1<>"����" then
       Response.Write "<script language=JavaScript>{alert('�����ǵ���,û��Ȩ�������������ս��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
rs.open "select ְλ,����,���� FROM �û� WHERE ����='" & to1 &"'",conn
zhiwei2=rs("ְλ")
guo2=rs("����")
pai2=rs("����")
 if pai2="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('�Է��ǹٸ�����,��������,��Զ��ȥ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if guo2="" or guo1="��" then
       Response.Write "<script language=JavaScript>{alert('�Է��޹��޽�,����ô����ɱ��ôһ�����أ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if zhiwei2<>"����" then
       Response.Write "<script language=JavaScript>{alert('�Է����ǹ���,û��Ȩ������Է����Ҳ����ս��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
  rs.close
       sql="SELECT d,bh FROM guo WHERE  g='" & guo1 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('������Ǳ������ˣ���ʲô���ǣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  enemy1=rs("d")
  protect1=rs("bh")
    if InStr("|" & enemy1 & "|", "|" & guo2& "|")>0 then
Response.Write "<script language=JavaScript>{alert('�����Ѿ���ս��,����Ҫ����ս�˰�.');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
       sql="SELECT d,bh FROM guo WHERE  g='" & guo2 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('�Է����������ǹ������ˣ����ǷŹ���');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
 enemy2=rs("d")
  protect2=rs("bh")
   if InStr("|" & enemy2 & "|", "|" & guo1& "|")>0 then
Response.Write "<script language=JavaScript>{alert('�Է������Ѿ����㿪ս��,����Ҫ����ս�˰�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       enemys1=enemy1&guo2&"|"
       enemys2=enemy2&guo1&"|"
   conn.execute"update guo set d='"&enemys1&"' where g='" & guo1 & "'"
    conn.execute"update guo set d='"&enemys2&"' where g='" & guo2 & "'"
guoopen=aqjh_name & "</font></b>��<b><font color=black>" & to1 & "</font></b>˵�� �����ʱ��������۱�������Ӣ����,�������һƴ����,�е����ľ����������ǹ�����û�����ʵ��������." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>