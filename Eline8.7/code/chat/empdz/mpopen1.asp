<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'������͡�wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Խ���������ͣ�');}</script>"
	Response.End
end if
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
says="<font color=green>��������͡�</font><font color=" & saycolor & ">"+mpopen1(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�������
function mpopen1(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
if Weekday(date())<>6 or (Hour(time())<21 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('��ʾ�����ɴ�սֻ����������21-22����У�����ֻ���������ϲ����ʸ�����ǰ������׼���ã�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
conn.open connstr
sql="select ���,���� from �û� where ����='" &sjjh_name&"'"
set rs=conn.execute(sql)
shenfen1=rs("���")
men1=rs("����")
 if men1="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('���ǹٸ����ˣ���ô����������������ͣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if men1="" or men1="��" then
       Response.Write "<script language=JavaScript>{alert('���������ɣ���ô�����������ѽ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if shenfen1<>"����" then
       Response.Write "<script language=JavaScript>{alert('���������ţ�û��Ȩ�������ɲ���������ͣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

sql="select ���,���� from �û� where ����='" &to1&"'"
set rs=conn.execute(sql)
shenfen2=rs("���")
men2=rs("����")
 if men2="�ٸ�" then
       Response.Write "<script language=JavaScript>{alert('�Է��ǹٸ����ˣ���ô�����и�����ս��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if men2="" or men1="��" then
       Response.Write "<script language=JavaScript>{alert('�Է���������,��ô�����������ѽ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
if shenfen2<>"����" then
       Response.Write "<script language=JavaScript>{alert('�Է��������ţ�û��Ȩ������Է����ɲ���������ͣ�');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
  rs.close
       sql="SELECT q,h,f FROM p WHERE  a='" & men1 & "'"
       rs.open sql,conn,1,1
       if rs.eof or rs.bof then 
       Response.Write "<script language=JavaScript>{alert('�Է���������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if
  shidi1=rs("q")
  baohu1=rs("f")
  kunjin1=rs("h")
  if baohu1=1 then
       Response.Write "<script language=JavaScript>{alert('���������ܹٸ���������ô���˼����');}</script>"
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

    if InStr("|" & shidi1 & "|", "|" & men2& "|")=0 then
Response.Write "<script language=JavaScript>{alert('�����Ѿ���ͣ�����Ҫ�ٴ������.');}</script>"
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
  if baohu2=1 then
       Response.Write "<script language=JavaScript>{alert('���������ܹٸ��������㲻��Ҫ���������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if

   if InStr("|" & shidi2 & "|", "|" & men1& "|")=0 then
Response.Write "<script language=JavaScript>{alert('�����Ѿ���ͣ�����Ҫ�ٴ������.');}</script>"
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
if application("sjjh_mp")<>""  then
dui=split(application("sjjh_mp"),"|")
if DateDiff("s",dui(1),sj2)<60 then
Response.Write "<script Language=Javascript>alert('��ǰ������ʱ�仹��" & 60-DateDiff("s",dui(1),sj2) & "��!');</script>"
response.end
else
Application.Lock
application("sjjh_mp")=to1&"|"&sj2&"|"&fn1&"|"&sjjh_name&"|"&men1&"|"&men2
Application.UnLock
mpopen1="["&sjjh_name&"]��<b><font color=black>["&to1&"]</font></b>˵�����������µ��Ӱ��ͺ�����������ɺ����ദ�����ݻ�����ϣ�������д���<input type=button value='Ը��' onclick=""javascript:window.open('empdz/mpopen1ok.asp?id="&myid&"&xiaozhu="&xiamoney&"&to1="&men2&"&yn=1&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='�ܾ�' onclick=""javascript:window.open('empdz/mpopen1ok.asp?id="&men1&"&xiaozhu="&xiamoney&"&to1="&men2&"&yn=0&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"">"
end if
else
Application.Lock
application("sjjh_mp")=to1&"|"&sj2&"|"&fn1&"|"&sjjh_name&"|"&men1&"|"&men2

Application.UnLock

mpopen1="["&sjjh_name&"]</font></b>��<b><font color=black>["&to1&"]</font></b>˵�����������µ��Ӱ��ͺ�����������ɺ����ദ�����ݻ�����ϣ�������д���<input type=button value='Ը��' onclick=""javascript:window.open('empdz/mpopen1ok.asp?id="&myid&"&xiaozhu="&xiamoney&"&to1="&men2&"&yn=1&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""button223""><input type=button value='�ܾ�' onclick=""javascript:window.open('empdz/mpopen1ok.asp?id="&men1&"&xiaozhu="&xiamoney&"&to1="&men2&"&yn=0&name="&to1&"','a','width=10,height=10')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" >"
end if
end function

%>