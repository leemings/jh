<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="config.asp"-->
<!--#include file="pass.asp"-->
<%Response.Expires=0
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
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
'if mid(server_v1,8,len(server_v2))<>server_v2 then
'        response.write "�Ͻ�ʹ���ݷֹ��ߣ�"
'        response.end
'end if
if Application("sjjh_room")="" then Response.Redirect "error.asp?id=000"
sername=Request.ServerVariables("SERVER_NAME")
ip=Request.ServerVariables("LOCAL_ADDR")
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
if InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE")=0 then Response.Redirect "error.asp?id=010"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if sjjh_disproxy="1" and (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "error.asp?id=011"
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
if Application("sjjh_closedoor")="1" then
Response.Write "<script Language=Javascript>alert('��ӭ���Ĺ��٣�\n���������������̫�ȶ�������վ��������ʱ�ر������ҵĵ�¼����\n���ڲ��ܽ��е�¼�����Щʱ��������\n���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ���� ����\n��л���Ա�������֧�ֺͺ񰮣�\n���������Ĳ��㣬���б�Ǹ��\nSEE YOU\n\n2003-11-10');window.parent.close();</script>"
Response.End
end if
name=Trim(Request.Form("name"))
password=Trim(Request.Form("pass"))
if instr(sjjh_disloginname,name)<>0 then Response.Redirect "error.asp?id=131"
if session("sjjh_name")<>"" then
	Session.Abandon
	Response.End 
end if

myid=Request.form("id")
if myid="" then
	Response.Write "<script Language=Javascript>alert('��ʲô��Ц����������ID����������');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,���� from �û� where id="& myid &" and ����='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('����������������Ҳ���!\n��鿴����ID�����������룡');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id=rs("id")
		
if session("sjjh_name")<>"" and session("sjjh_name")=name then 
	Response.Redirect "desktop/"
	Response.End 
end if
sjjh_roominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(sjjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("sjjh_useronlinename"&i))," "&LCase(name)&" ")=0 then ydl=0
	if ydl=1 then 
		Session.Abandon
		Response.Redirect "error.asp?id=140"
		Response.End 
	end if
next 
if session("sjjh_name")<>"" then 
	Response.Redirect "desktop/"
	Response.End 
end if
if len(name)>5  then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&name&"]�������ȣ����Ϊ10���ַ���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if password="" or name="" then Response.Redirect "error.asp?id=128"
if server.URLEncode(password)<>password then Response.Redirect "error.asp?id=121"
if chuser(name) then Response.Redirect "error.asp?id=120"
badword="����,�ڿ�,�ٿ�,��ʦ,�侫,��,ȥ��,����,��,��ʺ,����,����,����,��,ɫ,��,����,ɫ��,������,����,��,����,����,���,����,����,����,����,����,�쵰,�뵰,����,����,����,����,����,����,غ��,��ȥ,��Ů,����,����,����,�й�,���������,���������,���������,��Ƥ,��ͷ,��,��,è,��,�P,��,�H,����,��,��,����,����,����,����,����,����,ʮ����,��ү,������,���ϰ�,������,������,���,�Ҷ�,����,��,����,����,�Ҹ�,�Ҳ�,��ĸ,�Ҹ�,ȥ��,����,���׹�,����,��,��,��,ץ,��,��,��,ү,��,��,��,��,Ь,��,�,��,��,��,��,��,ʺ,��,��,��,��,��,�,��,��,��,��,��,��,ͷ,��,�P,��,�H,��,��,��,��,��,��,��,��,��ϯ,��,�ҿ�,����,��־,������,�ٸ�,����,����,���,��Ժ,����,˾��,��"
bad=split(badword,",")
for i=0 to ubound(bad)-1
	if InStr(LCase(name),bad(i))<>0 then Response.Redirect "error.asp?id=131"
next
dieip=sjjh_dieip
ipk=split(userip,".",-1)
if Instr(dieip,"*.*.*.*")<>0 or Instr(dieip,ipk(0)&".*.*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&".*.*")<>0 or Instr(dieip,ipk(0)&"."&ipk(1)&"."&ipk(2)&".*")<>0 or Instr(dieip,userip)<>0 then Response.Redirect "error.asp?id=111"
iplocktime=int(Application("sjjh_iplocktime"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
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

sjjh_name=rs("����")
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
'��ͨ��������
if rs("����")<-1000 or rs("����")<-1000 then
	conn.execute("update �û� set ͨ��=True,����=false where ����='"&sjjh_name&"'")
end if
if DateDiff("d",date(),rs("��Ա����"))<=0 and rs("��Ա")=True then
	conn.Execute "update �û� set ��Ա=False where ����='"&sjjh_name&"'"
	Response.Write "<script Language=Javascript>alert('��ʾ�����[�ݵ��ƻ�Ա]�Ѿ�����!');</script>"
end if

sjjh_jhdj=rs("�ȼ�")
mygj=rs("�ȼ�")*sjjh_gjsx+4500 
myfy=rs("�ȼ�")*sjjh_fysx+3000
if rs("����")>mygj then
	conn.execute "update �û� set ����="& mygj &" where ����='"&sjjh_name&"'"
end if
if rs("����")>myfy then
	conn.execute "update �û� set ����="& myfy &" where  ����='"&sjjh_name&"'"
end if

if sjjh_jhdj=60 and rs("��Ա�ȼ�")=1 then
conn.Execute "update �û� set ��Ա�ȼ�=2,��Ա����='"&sjj&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ��ȼ���60����2����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if sjjh_jhdj=90 and rs("��Ա�ȼ�")=2 then
conn.Execute "update �û� set ��Ա�ȼ�=3,��Ա����='"&sjj&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ��ȼ���90����3����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if sjjh_jhdj=90 and rs("��Ա�ȼ�")=1 then
conn.Execute "update �û� set ��Ա�ȼ�=2,��Ա����='"&sjj&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ�㣬��2����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if sjjh_jhdj=120 and rs("��Ա�ȼ�")=3 then
conn.Execute "update �û� set ��Ա�ȼ�=4,��Ա����='"&huiqi&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ��ȼ���120����4����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if sjjh_jhdj=120 and rs("��Ա�ȼ�")=2 then
conn.Execute "update �û� set ��Ա�ȼ�=3,��Ա����='"&huiqi&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ�㣬��3����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if sjjh_jhdj=120 and rs("��Ա�ȼ�")=1 then
conn.Execute "update �û� set ��Ա�ȼ�=2,��Ա����='"&huiqi&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ�㣬��2����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if sjjh_jhdj=300 and rs("��Ա�ȼ�")=4 then
conn.Execute "update �û� set ��Ա�ȼ�=5,��Ա����='"&huiqi&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ��ȼ���300����5����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if sjjh_jhdj=300 and rs("��Ա�ȼ�")=3 then
conn.Execute "update �û� set ��Ա�ȼ�=4,��Ա����='"&huiqi&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ�㣬��4����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if sjjh_jhdj=300 and rs("��Ա�ȼ�")=2 then
conn.Execute "update �û� set ��Ա�ȼ�=3,��Ա����='"&huiqi&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ�㣬��3����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if sjjh_jhdj=300 and rs("��Ա�ȼ�")=1 then
conn.Execute "update �û� set ��Ա�ȼ�=2,��Ա����='"&huiqi&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ�㣬��2����Ա�ˣ��ǵö������������ѽ ^_^');</script>"
end if
if rs("��Ա�ȼ�")=0 and rs("��Ա")=false then
conn.Execute "update �û� set ��Ա�ȼ�=1,��Ա����='"&sjjj&"' where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('��ʾ��վ���ٴ�����1����Ա�����£�ҪŬ���� ^_^');</script>"
end if

yin=rs("����")
cun=rs("���")
if (yin+cun)>2000000000 then
conn.Execute "update �û� set ����=100000000,���=900000000 where ����='"&sjjh_name&"'"
Response.Write "<script Language=Javascript>alert('ϵͳ���ѣ��������������ܺͳ���20�ڣ�ϵͳ�Զ�Ϊ����Ϊ10�ڣ��Ժ�Ҫע���ˣ� ^_^');</script>"
end if

sjjh_hy=rs("��Ա�ȼ�")
sjjh_hydate=rs("��Ա����")
hygq=DateDiff("d",date(),sjjh_hydate)
if sjjh_hy>0 and hygq<5 then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ļ�Ա����"&hygq&"�죬���㾡����վ����ϵ��');</script>"
end if
if sjjh_hy>0 and hygq<=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ļ�Ա�Ѿ�������ĸ���ֵ��ָ����ǻ�Ա״̬��');</script>"
	hydj=sjjh_hy
	Select Case hydj
	case 1
		conn.Execute "update �û� set ��Ա�ȼ�=0 where ����='"&sjjh_name&"'"
	case 2
		conn.Execute "update �û� set ��Ա�ȼ�=0 where ����='"&sjjh_name&"'"
	case 3
		conn.Execute "update �û� set ��Ա�ȼ�=0 where ����='"&sjjh_name&"'"
	case 4
		conn.Execute "update �û� set ��Ա�ȼ�=0,grade=1,����='����',���='����' where ����='"&sjjh_name&"'"
	end select
end if
zt=trim(rs("״̬"))
dldate=rs("��¼")
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
sjjh_hy=rs("��Ա�ȼ�")
prevtime=CDate(rs("lasttime"))
value=clng(rs("allvalue"))
dengji=int(sqr(value/50))
if DateDiff("m",prevtime,sj)<>0 then
	sqlstr="update �û� set mvalue=0,�ȼ�="& dengji &",times=times+1,lasttime='"&sj&"',lastip='"&userip&"' where ����='"&sjjh_name&"'"
else
	sqlstr="update �û� set �ȼ�="& dengji &",times=times+1,lasttime='"&sj&"',lastip='"&userip&"' where ����='"&sjjh_name&"'"
end if
conn.execute( sqlstr )
Session("sjjh_id")=rs("id")
Session("sjjh_name")=rs("����")
Session("sjjh_grade")=rs("grade")
Session("sjjh_jhdj")=rs("�ȼ�")
session("nowinroom")=0
if rs("����")=Application("sjjh_user") then
session("b28")="ok"
end if
Session.Timeout=300
response.cookies("yxjh").expires=date()+10 
response.cookies("yxjh")("sjjh_sid")=session.sessionid
if Session("sjjh_grade")>=10 and instr(Application("sjjh_admin"),Session("sjjh_name"))=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	session.Abandon
	Response.Redirect "error.asp?id=478"
	response.end
end if
if day(dldate)<>day(now()) or month(dldate)<>month(now()) or year(dldate)<>year(now()) then
	conn.execute("update �û� set ɱ����=0 where ����='" & sjjh_name & "'")
end if
conn.execute("update �û� set ״̬='����',�¼�ԭ��='��',��¼=now() where ����='"&sjjh_name&"'")
hy_sx=0
if rs("��Ա")=True and DateDiff("h",prevtime,sj)<24 then
	hy_sx=1
end if
if sjjh_sx=1 and hy_sx<>1 then
	wgj=(dengji*sjjh_wgsx+3800)+rs("�书��")
	conn.execute("update �û� set �书="& wgj &" where �书>"& wgj &" and ����='"&sjjh_name&"'")
	nlj=(dengji*sjjh_nlsx+2000)+rs("������")
	conn.execute("update �û� set ����="& nlj &" where ����>"& nlj &" and ����='"&sjjh_name&"'")
	tlj=(dengji*sjjh_tlsx+5260)+rs("������")
	conn.execute("update �û� set ����="& tlj &" where ����>"& tlj &" and ����='"&sjjh_name&"'")
end if
'����������
conn.execute("update sm set b=b+1 where a='������'")

'�ȼ���û��Ƿ��¼����������̳���Ƿ��н����û������ϣ����û�оͼ���
   dim bbsconn,rs1,rsbbs,question,answer,passjh,readme,myboardid
   db="bbs/data/Eline_bbs_6.3.0.asp"
   Set bbsconn = Server.CreateObject("ADODB.Connection")
   Set rs1=Server.CreateObject("ADODB.RecordSet")
   Set rsbbs=Server.CreateObject("ADODB.RecordSet")
   connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
   '�����ķ��������ý��ϰ汾Access�����������������ӷ���
   'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(db)
   bbsconn.open connstr

   sql="select * from [user] where username='"&session("sjjh_name") &"'"
   rsbbs.open sql,bbsconn,1,3
   if rsbbs.eof then
       rsbbs.addnew
       rsbbs("Article")=0
       set rs1=bbsconn.execute("select usertitle,titlepic from usertitle where not minarticle=-1 order by minarticle")
       rsbbs("userclass")=rs1(0)
       rsbbs("titlepic")=rs1(1)
       rs1.close
  
       rsbbs("showRe")=0
       rsbbs("logins")=1

	randomize timer
   	passjh=int(rnd*8998999)+1000
   	question=int(rnd*8998999)+1000
   	answer=int(rnd*8998999)+1000

   	rsbbs("username")=rs("����")

   	rsbbs("userpassword")=mid(md5(passjh),1,10)
        rsbbs("userWealth")=600
   	rsbbs("useremail")=rs("����")
   	rsbbs("oicq")=rs("oicq")
   	rsbbs("quesion")=mid(md5(question),1,10)
   	rsbbs("answer")=mid(md5(answer),1,10)
   end if
   '�޸���ݣ�����ǹٸ����ž�����Ϊ1Ϊmaster����������������ž�������3 boardmaster
   rsbbs("usergroupid")=4
   if rs("����")="�ٸ�" and rs("���")="����" then  Rsbbs("usergroupid")=1
   if rs("����")<>"�ٸ�" and rs("���")="����" then  Rsbbs("usergroupid")=3
   rsbbs("title")=rs("���")
   rsbbs("usergroup")=rs("����")
   if rs("�Ա�")="Ů"  then rsbbs("sex")=0 else rsbbs("sex")=1 end if
   Rsbbs("lastlogin")=NOW()
   rsbbs.update
   rsbbs.close
   '�ټ��ٸ������ǲ�admin���У����û�оͼ���

   if rs("����")="�ٸ�" and rs("���")="����" then
	sql="select * from admin where username='"&session("sjjh_name")&"'"
   	rsbbs.open sql,bbsconn,1,3
   	if rsbbs.eof then
   	   rsbbs.addnew
     	   rsbbs("username")=session("sjjh_name")
   	end if
     	rsbbs("password")=rs("����")
     	rsbbs("flag")="01, 02, 03, 04, 11, 12, 13, 14, 15, 16, 17, 21, 22, 23, 24, 25, 26, 27, 28, 31, 32, 33, 34, 35, 41, 42, 43, 44, 51, 52, 53, 54, 55, 61, 62, 63, 64, 71, 72, 73, 74, 75"
   	rsbbs("adduser")=session("sjjh_name")
   	rsbbs.update
   	rsbbs.close
   end if
   'ȡ�û�������дcookieΪ��̳�û���
        dim cookies_path_s,cookies_path_d,cookies_path,usercookies
   	'�жϸ���cookiesĿ¼
	cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
	cookies_path_d=ubound(cookies_path_s)
	cookies_path="/"
	for i=1 to cookies_path_d-1
		if not (cookies_path_s(i)="upload" or cookies_path_s(i)="admin") then 	cookies_path=cookies_path&cookies_path_s(i)&"/"
		
	next
	if cookiepath<>cookies_path then
		cookiepath=cookies_path
		bbsconn.execute("update config set cookiepath='"&cookiepath&"'")
	end if

        sql="select userid,userpassword,userclass from [user] where username='"&session("sjjh_name")&"'"
        rsbbs.open sql,bbsconn,1,1
	Response.Cookies("aspsky")("username") = session("sjjh_name")
	Response.Cookies("aspsky")("userid") = rsbbs("userid")
	Response.Cookies("aspsky")("password") = rsbbs("userpassword")
	Response.Cookies("aspsky")("userclass") = rsbbs("userclass")
	Response.Cookies("aspsky")("userhidden") = 2
	Response.Cookies("aspsky")("usercookies")="0"
	rem ���ͼƬ�ϴ���������
	response.cookies("upNum")=0
	Response.Cookies("aspsky").path=cookiepath
	
	rsbbs.close
   '��������Ƿ���groupname���У����û�оͼ���
   if rs("����")<>"�ٸ�" and rs("���")="����" then
     sql="select * from groupname where groupname='"&rs("����")&"'"
     rsbbs.open sql,bbsconn,1,3
     if rsbbs.eof then
       rsbbs.addnew
       rsbbs("groupname")=rs("����")
       rsbbs.update
     end if
     rsbbs.close
     '������������Ƿ����Լ�����̳��
     sql="select * from board where boardtype='�� "&rs("����")&" ��'"
     rsbbs.open sql,bbsconn,1,3
     if rsbbs.eof then
       'ȡ���Ű���
       'ȡ���ɵ�����
       set rs1=conn.execute("select d from p where a='"&rs("����")&"'")
       readme=rs1(0)
       rs1.close
       set rs1=bbsconn.execute("select max(boardid)+1 from board")
       myboardid=rs1(0)
       rs1.close
       set rs1=bbsconn.execute("select boardid,rootid,orders from board where boardtype='���Ű���'")
       if not rs1.eof then
          rsbbs.addnew
          rsbbs("boardid")=myboardid
          rsbbs("boardtype")="�� "&rs("����")&" ��"
          rsbbs("parentid")=rs1(0)
          rsbbs("parentstr")=rs1(0)
          rsbbs("depth")=1
          rsbbs("rootid")=rs1(1)
          rsbbs("child")=0
          rsbbs("orders")=rs1("orders")+1
          rsbbs("readme")=readme
          rsbbs("boardmaster")=""
          rsbbs("lastbbsnum")=0
          rsbbs("lasttopicnum")=0
          rsbbs("indeximg")=""
          rsbbs("todaynum")=0
          rsbbs("boarduser")=""
          rsbbs("lastpost")="$0$"&date()&" "&now()&"$$$$$"
          rsbbs("sid")=1
          rsbbs("board_setting")="0,0,0,0,1,0,1,1,1,1,1,1,1,1,1,1,16240,3,300,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,0|24,1,0,300,20,10,9,12,1,10,10,0,0,0,0,1,5,0,1,4,0,0,0,0,0,0,0,0,0"
          rsbbs.update
       end if
       rs1.close
       
       rsbbs.close
     end if
     
   end if
   set rsbbs=nothing
   set rs1=nothing

'��Ա��ʾ
rs.close
rs.open "SELECT * FROM sm where a='����'",conn,2,2
banner=rs("c")
mygg=rs("d")
if sjjh_hy>0 then
	Response.Write "<script Language=Javascript>alert('"&mygg&"');</script>"
else
	banners=Split(Trim(banner),";",-1)
	total=UBound(banners)
	randomize timer
	x=int(rnd*(total+1))
	Response.Write "<script Language=Javascript>alert('"&banners(x)&"');</script>"
	erase banners
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>chatwin=window.open('chat/jhchat.asp','sjjh','left=0,top=0,status=no,scrollbars=no,resizable=no');chatwin.resizeTo(screen.availWidth,screen.availHeight);</script>"
if Request.form("txwz")=0 then
Response.Write "<script Language=Javascript>location.href = 'skin.asp';</script>"
end if
if Request.form("txwz")=1 then
Response.Write "<script Language=Javascript>location.href = 'mh.asp';</script>"
end if
if Request.form("txwz")=2 then
Response.Write "<script Language=Javascript>location.href = 'cww.asp';</script>"
end if
if Request.form("txwz")=3 then
Response.Write "<script Language=Javascript>location.href = 'gs.asp';</script>"
end if
if Request.form("txwz")=4 then
Response.Write "<script Language=Javascript>location.href = 'szwd.asp';</script>"
end if
if Request.form("txwz")=5 then
Response.Write "<script Language=Javascript>location.href = 'sm.asp';</script>"
end if
if Request.form("txwz")=6 then
Response.Write "<script Language=Javascript>location.href = 'lr.asp';</script>"
end if
if Request.form("txwz")=7 then
Response.Write "<script Language=Javascript>location.href = 'qs.asp';</script>"
end if
if Request.form("txwz")=8 then
Response.Write "<script Language=Javascript>location.href = 'th.asp';</script>"
end if
if Request.form("txwz")=9 then
Response.Write "<script Language=Javascript>location.href = 'ppb.asp';</script>"
end if
if Request.form("txwz")=10 then
Response.Write "<script Language=Javascript>location.href = 'neweline.asp';</script>"
end if
if Request.form("txwz")=11 then
Response.Write "<script Language=Javascript>location.href = 'desktop/';</script>"
else
Response.Write "<script Language=Javascript>location.href = 'mh.asp';</script>"
end if
%>