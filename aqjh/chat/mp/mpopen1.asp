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
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Խ���������ͣ�');}</script>"
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
says="<font color=green>��������͡�</font><font color=" & saycolor & ">"+mpopen1(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�������
function mpopen1(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
sql="select ���,���� from �û� where ����='" &aqjh_name&"'"
set rs=conn.execute(sql)
life1=rs("���")
pai1=rs("����")
 if pai1="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('���ǹٸ����ˣ���ô����������������ͣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if pai1="" or pai1="��" then
       Response.Write "<script language=JavaScript>{alert('����������,��ô�����������ѽ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if life1<>"����" then
       Response.Write "<script language=JavaScript>{alert('����������,û��Ȩ�������ɲ���������ͣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

sql="select ���,���� from �û� where ����='" &to1&"'"
set rs=conn.execute(sql)
life2=rs("���")
pai2=rs("����")
 if pai2="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('�Է��ǹٸ�����,��ô�����и�����ս��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if pai2="" or pai1="��" then
       Response.Write "<script language=JavaScript>{alert('�Է���������,��ô�����������ѽ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if life2<>"����" then
       Response.Write "<script language=JavaScript>{alert('�Է���������,û��Ȩ������Է����ɲ���������ͣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
  rs.close
       sql="SELECT q,h,f FROM p WHERE  a='" & pai1 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('�Է���������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  enemy1=rs("q")
  protect1=rs("f")
  kunjin1=rs("h")
  if protect1=1 then
       Response.Write "<script language=JavaScript>{alert('���������ܹٸ�����,��ô���˼����');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

  if kunjin1<=xiamoney then
       Response.Write "<script language=JavaScript>{alert('��������ô��Ŀ����?');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

    if InStr("|" & enemy1 & "|", "|" & pai2& "|")=0 then
Response.Write "<script language=JavaScript>{alert('�����Ѿ����,����Ҫ�ٴ������.');}</script>"
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
  if protect2=1 then
       Response.Write "<script language=JavaScript>{alert('���������ܹٸ�����,�㲻��Ҫ���������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

   if InStr("|" & enemy2 & "|", "|" & pai1& "|")=0 then
Response.Write "<script language=JavaScript>{alert('�����Ѿ����,����Ҫ�ٴ������.');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
       n=Year(date())
     y=Month(date())
     r=Day(date())
     s=Hour(time())
     f=Minute(time())
     m=Second(time())
     if len(y)=1 then y="0" & y
     if len(r)=1 then r="0" & r
     if len(s)=1 then s="0" & s
     if len(f)=1 then f="0" & f
     if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
if application("aqjh_mp")<>""  then
dui=split(application("aqjh_mp"),"|")
if DateDiff("s",dui(1),sj2)<60 then
Response.Write "<script Language=Javascript>alert('��ǰ������ʱ�仹��" & 60-DateDiff("s",dui(1),sj2) & "��!');</script>"
response.end
else
Application.Lock
application("aqjh_mp")=to1&"|"&sj2&"|"&fn1&"|"&aqjh_name&"|"&pai1&"|"&pai2
Application.UnLock
mpopen1="["&aqjh_name&"]��<b><font color=black>["&to1&"]</font></b>˵�����������µ��Ӱ��ͺ���,������ɺ����ദ,���ݻ���.ϣ�������д���<input type=button value='Ը��' onclick=""javascript:window.open('mp/mpopen1ok.asp?id="&myid&"&xiaozhu="&xiamoney&"&to1="&pai2&"&yn=1&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='�ܾ�' onclick=""javascript:window.open('mp/mpopen1ok.asp?id="&pai1&"&xiaozhu="&xiamoney&"&to1="&pai2&"&yn=0&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"">"
end if
else
Application.Lock
application("aqjh_mp")=to1&"|"&sj2&"|"&fn1&"|"&aqjh_name&"|"&pai1&"|"&pai2

Application.UnLock

mpopen1="["&aqjh_name&"]</font></b>��<b><font color=black>["&to1&"]</font></b>˵�����������µ��Ӱ��ͺ���,������ɺ����ദ,���ݻ���.ϣ�������д���<input type=button value='Ը��' onclick=""javascript:window.open('mp/mpopen1ok.asp?id="&myid&"&xiaozhu="&xiamoney&"&to1="&pai2&"&yn=1&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='�ܾ�' onclick=""javascript:window.open('mp/mpopen1ok.asp?id="&pai1&"&xiaozhu="&xiamoney&"&to1="&pai2&"&yn=0&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" >"
end if
end function

%>