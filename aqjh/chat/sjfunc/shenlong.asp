<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%
aqjh_roominfo=split(Application("aqjh_room"),";")
nowinroom=session("nowinroom")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
onlinekill=Application("aqjh_onlinekill")
'if onlinenow<onlinekill and chatinfo(0)<>"" then
	'Response.Write "<script language=JavaScript>{alert('��ʾ:�������ߵ���"&onlinekill&"�˲��ö��䣡');}</script>"
	'Response.End
'end if
next
%>
<%'��������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")


if chatinfo(5)<>0 or nowinroom=3 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Է��У�');}</script>"
	Response.End
end if
f=Minute(time()) 
if f>30 and nowinroom<>2 and nowinroom<>4 then
        Response.Write "<script language=JavaScript>{alert('��ʾ������PK���ʱ��ΪÿСʱǰ30�֣������ȥ��ս��');}</script>"
	response.end
end if
if aqjh_jhdj<31 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ31�����ϲſ��Բ�����');}</script>"
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
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
   pk="������������"

says="<font color=red>"&pk&"<font color=" & saycolor & ">"+shen(mid(says,i+1),towho)+"</font>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)


function shen(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('������ʲô���뵷���𣿣�');}</script>"
Response.End
exit function
end if


if to1=aqjh_name or to1="���" then
	Response.Write "<script language=JavaScript>{alert('��ѡ����ȷ�Ķ������');}</script>"
       Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")

sql="select * from myanimal where username='"&aqjh_name&"' and animalname='"&trim(fn1)&"'  and rest=0"
rs.open sql,conn,1,1
      if rs.bof or rs.eof then
        Response.Write "<script language=JavaScript>{alert('��û��["&fn1&"]��������!����������Ŀǰ�����ٸ������ˣ�');}</script>"
        rs.close
	 set rs=nothing	
	 conn.close
	 set conn=nothing
        Response.End
       end if
gong1=rs("attack")
sm=rs("sm")
lx=rs("lei")
rs.close
''''''''''''''''''''''''''''''''''''''''''''''


sql="select ����,����,����ʱ�� from �û� where ����='" & aqjh_name&"'"
	rs.open sql,conn,1,1
       mp=rs("����")
       baohu1=rs("����")
        reg=rs("����ʱ��")
      if baohu1=true then
	Response.Write "<script language=JavaScript>{alert('��Ŀǰ���ڱչ����������ǵ㽭���ºò��ã���');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       Response.End
       end if
    '   if aqjh_grade>=6 then
    '   Response.Write "<script language=JavaScript>{alert('���ǹٸ��ģ����ܶԽ������˽�������������');}</script>"
   '    rs.close
	'set rs=nothing	
	'conn.close
	'set conn=nothing
    '   Response.End
    '   end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<15 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        shen="##���㲻�ܶ�%%���й��������ոձ�����ɱ���������ȷŹ����ɣ�"
        exit function
end if

       rs.close


''''''''''''''''''''''''''''''''''''''''''''''''
       sql="select grade,����,����,����,����ʱ��,�ȼ�  from �û� where ����='" & to1 & "'"
       rs.open sql,conn,1,1
       ntnt=rs("�ȼ�")
       baohu2=rs("����")
       reg1=rs("����ʱ��")
       topai=rs("����")
       ti=rs("����")
       tograde=rs("grade")
       if baohu2=true then
	Response.Write "<script language=JavaScript>{alert('�Է����������У��㲻�ܽ���������������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if

sj=DateDiff("n",rs("����ʱ��"),now())
if sj<15 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�%%���й�������ոձ�����ɱ�����������ǵ�ɣ�"
        exit function
end if

       if ntnt<=30  then
       Response.Write "<script language=JavaScript>{alert('�Է����ǽ������֣�����ô����ѽ��');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
	end if

       if tograde>=6 then
       Response.Write "<script language=JavaScript>{alert('�Է����ǽ����еĹٸ��������ǲ��ǲ����ڽ��������');}</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
       Response.End
       end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''
       conn.execute ("update �û� set ����=����-"&gong1&" where ����='"&to1&"'")
       conn.execute ("update myanimal set attack=attack*9/10 where username='"&nickname&"' and animalname='"&trim(fn1)&"'")

			if ti-gong1<-1000 then
                    	       conn.execute"update �û� set ״̬='��' where ����='" & to1 & "'"
				conn.execute"update �û� set allvalue=allvalue+"  & ntnt*5 & ",����=����-10,����=����-10,ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & aqjh_name & "'"
				conn.execute("update myanimal set username='��',rest=1 where username='" &towho& "'")
call boot(to1,"�����������ߣ�"&aqjh_name&",["&menpai&"]"&fn1)
                           shen="["&aqjh_name&"]�ų���������<b><font color=black>"&fn1&"("&lx&")</font></b>,<b><font color=red>"&sm&"</font></b><bgsound src=READONLY/long.mid loop=1>��ɱ��["&to1&"]����<b><font color=red>"&gong1&"</font></b>,"&to1&"ѧ�ղ�����������ҧ����"&aqjh_name&"Ϊɱ������Լ���ʧ��10����º�10������,�õ�"  & ntnt*5 & "�����飬"&aqjh_name&"�������������½�1/10"
                             rs.close
                             set rs=nothing	
	                      conn.close
	                      set conn=nothing	
                             exit function
                    end if
                       shen="["&aqjh_name&"]�ų���������<b><font color=black>"&fn1&"("&lx&")</font></b>,<b><font color=red>"&sm&"</font></b><bgsound src=READONLY/long.mid loop=1>��ɱ��["&to1&"]����<b><font color=red>"&gong1&"</font></b>,"&aqjh_name&"�������������½�1/10!" 
                        rs.close
                        set rs=nothing	
	                 conn.close
                       set conn=nothing  
                           
        
end function
%>
</body>
</html>