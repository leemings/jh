<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="config.asp"-->
<!--#include file="const2.asp"-->
<!--#include file="const3.asp"-->
<!--#include file="pass.asp"-->
<!--#include file="showchat.asp"-->
<!--#include file="chk.asp"-->
<%Response.Expires=0
call Chkproxy()
if aqjh_myie=1 then call Chkmyie()
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
call chkpost()
if Application("aqjh_room")="" then Response.Redirect "error.asp?id=000"
sername=Request.ServerVariables("SERVER_NAME")
ip=Request.ServerVariables("LOCAL_ADDR")
ip="127.0.0.1"
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
'if InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE")=0 then Response.Redirect "error.asp?id=010"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if aqjh_disproxy=1 then  Chkproxy()
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
if y<10 then
sjj=n & "-" & y+3 & "-" & r
else
sjj=n+1 & "-" & y+3-12 & "-" & r
end if
if y<11 then
sjjj=n & "-" & y+2 & "-" & r
else
sjjj=n+1 & "-" & y+2-12 & "-" & r
end if
huiqi=n+1 & "-" & y & "-" & r
userip=Request.ServerVariables("REMOTE_ADDR")
name=Trim(Request.Form("name"))
password=Trim(Request.Form("pass"))
if Application("aqjh_closedoor")="1" and name<>Application("aqjh_user") then Response.Redirect "error.asp?id=100"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
next
if onlinenow>Application("aqjh_chat_maxpeople") then Response.Redirect "error.asp?id=101"
if instr(aqjh_disloginname,name)<>0 then Response.Redirect "error.asp?id=130"
if session("aqjh_name")<>"" then
	Session.Abandon
	Response.End 
end if		
if session("aqjh_name")<>"" and session("aqjh_name")=name then 
	Response.Redirect "welcome.asp"
	Response.End 
end if
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(name)&" ")=0 then ydl=0
	if ydl=1 then 
		Session.Abandon
		Response.Redirect "error.asp?id=140"
		Response.End 
	end if
next 
if session("aqjh_name")<>"" then 
	Response.Redirect "welcome.asp"
	Response.End 
end if
if len(name)>5  then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&name&"]�������ȣ����Ϊ10���ַ���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if password="" or name="" then Response.Redirect "error.asp?id=128"
if server.URLEncode(password)<>password then Response.Redirect "error.asp?id=121"
if chuser(name) then Response.Redirect "error.asp?id=120"
badword="����,�ڿ�,�ٿ�,��ʦ,�侫,��,ȥ��,����,��,��ʺ,����,����,����,��,����,����,�й�,ɫ,��,����,ɫ��,������,����,��,����,����,���,����,����,����,����,����,�쵰,�뵰,����,����,����,����,����,����,غ��,��ȥ,��Ů,����,����,���������,���������,���������,��Ƥ,��ͷ,��,��,è,��,�P,��,�H,����,��,��,����,����,����,����,����,����,ʮ����,��ү,������,���ϰ�,������,������,���,�Ҷ�,����,��,����,����,�Ҹ�,�Ҳ�,��ĸ,�Ҹ�,ȥ��,����,���׹�,����,��,��,��,ץ,��,��,����,ү,��,��,��,��,Ь,��,�,��,��,��,��,��,ʺ,��,��,��,��,��,�,��,��,��,��,��,��,ͷ,��,�P,��,�H,��,��,��,��,��,��,��,��,��ϯ,��,�ҿ�,����,��־,������,�ٸ�,����,����,���,��Ժ,����,˾��,��"
bad=split(badword,",")
for i=0 to ubound(bad)-1
	if InStr(LCase(name),bad(i))<>0 then Response.Redirect "error.asp?id=131"
next
dieip=aqjh_dieip
ipk=split(userip,".",-1)
if Instr(dieip,"*.*.*.*")<>0 or Instr(dieip,ipk(0)&".*.*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&".*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&"."&ipk(2)&".*")<>0 or Instr(dieip,userip)<>0 then Response.Redirect "error.asp?id=111"
iplocktime=int(Application("aqjh_iplocktime"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dcz=0
sql="SELECT a FROM i WHERE DateDiff('n',b,#" & sj & "#)>=" & iplocktime
rs.open sql,conn,1,1
if Not(rs.Eof and rs.Bof) then dcz=1
rs.close
if dcz=1 then
	conn.Execute "DELETE FROM i WHERE DateDiff('n',b,#" & sj & "#)>=" & iplocktime
end if
sql="SELECT a,b FROM i WHERE a='" & userip & "'"
rs.open sql,conn,1,1
if NOT(rs.Eof and rs.Bof) then
	lockdate=rs("b")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=110&lockdate=" & server.URLEncode(lockdate)
end if
rs.close
'ת�����룺
password=md5(password)
sql="SELECT * FROM �û� WHERE ����='"&name&"'"
rs.open sql,conn,2,2
if rs.Eof and rs.Bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=423"
	response.end
end if
if rs("����")<>password then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=141"
	response.end
end if
aqjh_name=rs("����")
aqjh_grade=rs("grade")
aqjh_jhdj=rs("�ȼ�")
'reglastkick=rs("lastkick")
sjyy=rs("�¼�ԭ��")
if rs("����")<-100 or rs("״̬")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=421&sjyy="&sjyy
	response.end
end if
'���������书�������޵Ĵ���
if rs("����")<-1000 then
	conn.execute("update �û� set ����=0 where ����='"&aqjh_name&"'")
end if
if rs("�书")<-1000 then
	conn.execute("update �û� set �书=0 where ����='"&aqjh_name&"'")
end if
if rs("������")<0 then
	conn.execute("update �û� set ������=0 where ����='"&aqjh_name&"'")
end if
if rs("������")<0 then
	conn.execute("update �û� set ������=0 where ����='"&aqjh_name&"'")
end if
'��ͨ��������
if rs("����")<-1000 or rs("����")<-1000 then
	conn.execute("update �û� set ͨ��=True,����=false where ����='"&aqjh_name&"'")
end if
'�Խ�Ǯ�������޵Ĵ���
mess1="<font color=red>"&aqjh_name&"</font>�ȵ���������ҡͷ���Ը����������"
if rs("����")>1000000000 then
	mess="�ٸ�ͻȻ���֣���������ô��Ǯ��[����]������ʮ�����ո�������˰����"
        conn.execute "update �û� set ����=1000000000 where  ����='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('��ʾ���ٸ�ͻȻ���֣���������ô��Ǯ��[����]������ʮ�����ո�������˰����');</script>"
	if aqjh_grade<10 then call showchat("<font color=ff0000>���ٸ���˰��</font>" & mess1 & mess)
end if
if rs("���")>1000000000 then
	mess="�ٸ�ͻȻ���֣���������ô��Ǯ��[���]������ʮ�����ո�������˰����"
        conn.execute "update �û� set ���=1000000000 where  ����='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('��ʾ���ٸ�ͻȻ���֣���������ô��Ǯ��[���]������ʮ�����ո�������˰����');</script>"
	if aqjh_grade<10 then call showchat("<font color=ff0000>���ٸ���˰��</font>" & mess1 & mess)
end if
if rs("���")>1000000000 then
	mess="�ٸ�ͻȻ���֣���������ô��Ǯ��[���]������ʮ�����ո�������˰����"
        conn.execute "update �û� set ���=1000000000 where  ����='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('��ʾ���ٸ�ͻȻ���֣���������ô��Ǯ��[���]������ʮ�����ո�������˰����');</script>"
	if aqjh_grade<10 then call showchat("<font color=ff0000>���ٸ���˰��</font>" & mess1 & mess)
end if
'�Խ�Ǯ�������޴������
userid=rs("id")
if DateDiff("d",date(),rs("��Ա����"))<=0 and rs("��Ա")=True then
	conn.Execute "update �û� set ��Ա=False where ����='"&aqjh_name&"'"
	Response.Write "<script Language=Javascript>alert('��ʾ�����[�ݵ��ƻ�Ա]�Ѿ�����!');</script>"
end if
aqjh_hy=rs("��Ա�ȼ�")
aqjh_hydate=rs("��Ա����")
hygq=DateDiff("d",date(),aqjh_hydate)
if aqjh_hy>0 and hygq<5 then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ļ�Ա����"&hygq&"�죬���㾡����վ����ϵ��');</script>"
end if
if aqjh_hy>0 and hygq<=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ļ�Ա�Ѿ�������ĸ���ֵ��ָ����ǻ�Ա״̬��');</script>"
	hydj=aqjh_hy
	conn.Execute "update �û� set ��Ա�ȼ�=0,ְҵ='�ɱ�' where ����='"&aqjh_name&"'"
end if
zt=trim(rs("״̬"))
dldate=rs("��¼")
denglucha=DateDiff("d",dldate,now())
if denglucha>10 and rs("grade")>5 and rs("grade")<9 then
conn.execute("update �û� set grade=1,����='����' where ����='" &aqjh_name& "'")
Response.Write "<script Language=javascript>alert('�����ֽ�������ʾ�����Ѿ�"& denglucha &"��δ��¼���ֽ�����\n6-8���ٸ�����10��δ��¼���Զ���ְ��');</script>"
end if
if zt="���" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=420&sjyy="&sjyy
	response.end
end if
if rs("��¼")>now() and (zt="��" or zt="��" or zt="��Ѩ" or zt="��") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	select case zt
		case "��"
				Response.Redirect "error.asp?id=420&sjyy="&sjyy
				response.end
		case "��"
				Response.Redirect "error.asp?id=422&sjyy="&sjyy
				response.end
		case "��Ѩ" 
				Response.Redirect "error.asp?id=480&sjyy="&sjyy
				response.end
		case "��" 
				Response.Redirect "error.asp?id=472"
				response.end
	end select
end if
if rs("�¼�ԭ��")<>"��" then
	Response.Write "<script Language=Javascript>alert('��������:"&rs("�¼�ԭ��")&"');</script>"
end if
aqjh_hy=rs("��Ա�ȼ�")
prevtime=CDate(rs("lasttime"))
value=clng(rs("allvalue"))
dengji=int(sqr(value/50))
ltime=rs("lasttime")
lip=rs("lastip")
if DateDiff("m",prevtime,sj)<>0 then
	sqlstr="update �û� set mvalue=0,�ȼ�="& dengji &",times=times+1,lasttime='"&sj&"',lastip='"&userip&"' where ����='"&aqjh_name&"'"
else
	sqlstr="update �û� set �ȼ�="& dengji &",times=times+1,lasttime='"&sj&"',lastip='"&userip&"' where ����='"&aqjh_name&"'"
end if
conn.execute( sqlstr )
Session("aqjh_name")=rs("����")
Session("aqjh_grade")=rs("grade")
Session("aqjh_jhdj")=rs("�ȼ�")
Session("aqjh_id")=rs("id")
session("nowinroom")=0
Session.Timeout=60
response.cookies("aqjh").expires=date()+10 
response.cookies("aqjh")("aqjh_sid")=session.sessionid
if Session("aqjh_grade")>=10 and instr(Application("aqjh_admin"),Session("aqjh_name"))=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=478"
	response.end
end if
if day(dldate)<>day(now()) or month(dldate)<>month(now()) or year(dldate)<>year(now()) then
	conn.execute("update �û� set ɱ����=0 where ����='" & aqjh_name & "'")
end if
conn.execute("update �û� set ״̬='����',�¼�ԭ��='��',��¼=now() where ����='"&aqjh_name&"'")
hy_sx=0
if rs("��Ա")=True and DateDiff("h",prevtime,sj)<24 then
	hy_sx=1
end if
mess1="<font color=red>"&aqjh_name&"</font>�ڼ��������һ��ͨ���ľ����澭�������������"
if aqjh_sx=1 and hy_sx<>1 then
	wgj=(dengji*aqjh_wgsx+3800)+rs("�书��")
        if rs("�书")>wgj then
        mess2="[�书]"
        messok=1
	conn.execute("update �û� set �书="& wgj &" where ����='"&aqjh_name&"'")
        end if
	nlj=(dengji*aqjh_nlsx+2000)+rs("������")
        if rs("����")>nlj then
        mess3="[����]"
        messok=1
	conn.execute("update �û� set ����="& nlj &" where ����='"&aqjh_name&"'")
        end if
	tlj=(dengji*aqjh_tlsx+5260)+rs("������")
        if rs("����")>tlj then
	conn.execute("update �û� set ����="& tlj &" where ����='"&aqjh_name&"'")
        mess4="[����]"
        messok=1
        end if
 if aqjh_grade<10 and messok=1 then
    call showchat("<font color=ff0000>���߻���ħ��</font>"&mess1&"��Ϊ<font color=red>"&mess2&mess3&mess4&"</font>�������弫�޲���߻���ħ�������õõ����˵ĵ㻯�ָ�����")
 end if
'�Թ����ͷ�������
 mygj=(rs("�ȼ�")+1)*aqjh_gjsx+4500+rs("������")
 myfy=(rs("�ȼ�")+1)*aqjh_fysx+3000+rs("������")
 if rs("����")>mygj then
	conn.execute "update �û� set ����="& mygj &" where ����='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('��ʾ�����Ĺ����Ѿ���������!ϵͳ�Ѿ�����!');</script>"
 end if
 if rs("����")>myfy then
	conn.execute "update �û� set ����="& myfy &" where  ����='"&aqjh_name&"'"
        Response.Write "<script Language=Javascript>alert('��ʾ�����ķ����Ѿ���������!ϵͳ�Ѿ�����!');</script>"
 end if
end if
'����������
conn.execute("update [count] set num=num+1 where count='������'")
   '=========================================================
   ' BbsXp 5.13b for ���ֽ���9.9���
   '=========================================================
   dim bbsconn,rs1,rsbbs,UserQuesion,UserAnswer,passjh,readme,myboardid,jhtx
   db="bbs/database/aqjh_bbs.asp"
   Set bbsconn = Server.CreateObject("ADODB.Connection")
   Set rs1=Server.CreateObject("ADODB.RecordSet")
   Set rsbbs=Server.CreateObject("ADODB.RecordSet")
   connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
   bbsconn.open connstr
   sql="select * from [User] where username='"& session("aqjh_name") &"'"
   rsbbs.open sql,bbsconn,1,3
   if rsbbs.eof then
       rsbbs.addnew
       rsbbs("username")=rs("����")
       rsbbs("UserInfo")="\�й�\\\\\\\\\\\\\"
       rsbbs("UserIM")="\\\\\"
       rsbbs("friend")="|"
   end if
   '�޸���ݣ�����ǹٸ����ž�����Ϊ����վ��
   if rs("����")="�ٸ�" and rs("���")="����" and rs("����")=Application("aqjh_user") then
	Rsbbs("membercode")=5
	bbsconn.execute("update [clubconfig] set adminpassword='"&ucase(rs("����"))&"'")
   end if
   rsbbs("honor")=rs("���")
   rsbbs("faction")=rs("����")
   rsbbs("usermail")=rs("����")
   if rs("�Ա�")="Ů" then rsbbs("userface")="images/face/49.gif" else rsbbs("userface")="images/face/28.gif"
   randomize timer
   pass1=int(rnd*8998999)+1000
   pass2=int(rnd*8998999)+1000
   rsbbs("question")=mid(md5(pass1),1,10)
   rsbbs("answer")=mid(md5(pass2),1,10)
   if rs("�Ա�")="Ů" then rsbbs("Sex")="female" else rsbbs("Sex")="male" end if
   Rsbbs("landtime")=NOW()
   rsbbs("userpass")=mid(rs("����"),1,10)
   rsbbs.update
   rsbbs.close
   'ȡ�û�������дcookieΪ��̳�û���
        sql="select * from [User] where username='"&session("aqjh_name")&"'"
        rsbbs.open sql,bbsconn,1,1
        Response.Cookies("username")=rsbbs("username")
        Response.Cookies("userpass")=rsbbs("userpass")
	Response.Cookies("eremite")="0"
        rsbbs.close    
'��̳д�����
'��Ա��ʾ
rs.close
rs.open "SELECT * FROM sm where a='����'",conn,2,2
banner=rs("c")
mygg=rs("d")
if aqjh_hy>0 then
	Response.Write "<script Language=Javascript>alert('"&mygg&"\n\n���ϴε�½��ʱ���ǣ�"&ltime&"\n���ϴε�½��IP�ǣ�"&lip&"');</script>"
else
	banners=Split(Trim(banner),";",-1)
	total=UBound(banners)
	randomize timer
	x=int(rnd*(total+1))
	Response.Write "<script Language=Javascript>alert('"&banners(x)&"\n\n���ϴε�½��ʱ���ǣ�"&ltime&"\n���ϴε�½��IP�ǣ�"&lip&"');</script>"
	erase banners
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
if aqjh_myie=1 then
 call Chkmyie()
else
 Response.Write "<script Language=Javascript>chatwin=window.open('chat/jhchat.asp','aqjh','left=0,top=0,status=no,scrollbars=no,resizable=no');chatwin.resizeTo(screen.availWidth,screen.availHeight);</script>"
 Response.Write "<script Language=Javascript>location.href = 'jh.HTM';</script>"
end if
%>