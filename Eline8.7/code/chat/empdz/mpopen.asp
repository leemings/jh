<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'������ս��wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����ʹ��������ս��');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if Weekday(date())<>6 or (Hour(time())<21 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('��ʾ�����ɴ�սֻ����������21-22����У���ս����ֻ�����Ų����ʸ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
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
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="���" or towho=application("sjjh_automanname") or towho=sjjh_name then
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
conn.open Application("sjjh_usermdb")
rs.open "select ���,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,3,3
shenfen1=rs("���")
men1=rs("����")
 if men1="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('���ǹٸ����ˣ���ô�ܸ��������ɽ�Թ�������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if men1="" or men1="��" then
       Response.Write "<script language=JavaScript>{alert('���������ɣ���ô�������ɿ�սѽ��ȥ�������ɰɣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if shenfen1<>"����" then
       Response.Write "<script language=JavaScript>{alert('���������ţ�û��Ȩ�������ɲ���������ս��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
rs.open "select ���,���� FROM �û� WHERE ����='" & to1 &"'",conn
shenfen2=rs("���")
men2=rs("����")
 if men2="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('�Է��ǹٸ����ˣ�������������Զ��ȥ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if men2="" or men1="��" then
       Response.Write "<script language=JavaScript>{alert('�Է��������ɣ���ô����������սѽ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if shenfen2<>"����" then
       Response.Write "<script language=JavaScript>{alert('�Է��������ţ�û��Ȩ������Է����ɲ���������ս��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
  rs.close
       sql="SELECT q,f FROM p WHERE  a='" & men1 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('���������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  shidi1=rs("q")
  baohu1=rs("f")
    if InStr("|" & shidi1 & "|", "|" & men2& "|")>0 then
Response.Write "<script language=JavaScript>{alert('�����Ѿ���ս�ˣ�����Ҫ����ս�˰�.');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
rs.close
       sql="SELECT q,f FROM p WHERE  a='" & men2 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('�Է���������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
 shidi2=rs("q")
  baohu2=rs("f")
   if InStr("|" & shidi2 & "|", "|" & men1& "|")>0 then
Response.Write "<script language=JavaScript>{alert('�Է������Ѿ����㿪ս�ˣ�����Ҫ����ս�˰�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       shidis1=shidi1&men2&"|"
       shidis2=shidi2&men1&"|"
   conn.execute"update p set q='"&shidis1&"' where a='" & men1 & "'"
    conn.execute"update p set q='"&shidis2&"' where a='" & men2 & "'"
mpopen=sjjh_name & "</font></b>��<b><font color=black>" & to1 & "</font></b>˵�����������µ��ӷ���ս����������������ս����ɱ���һͳ������ҵ��" 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing

       
end function 
%>
