<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
un=session("yx8_mhjh_username")
grade=session("yx8_mhjh_usergrade")
chatroomsn=session("yx8_mhjh_userchatroomsn")
if un="" then Response.Redirect "../error.asp?id=016"
if grade<10 then
Response.Write "<script language=javascript>{alert('�Ĳ���Ҫ10������\n��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chatroom")=0 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if 
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "��ֹվ���ύ��"
        response.end
end if	
mg=abs(int(clng(Request.form("money"))))
idd=Request.form("id")
abcc=mid(mg,6)
for t=1 to len(abcc)
abc=mid(abcc,t,1)
if abc<>"0" and abc<>"1" and abc<>"2" and abc<>"3" and abc<>"4" and abc<>"5" and abc<>"6" and abc<>"7" and abc<>"8" and abc<>"9" then
	Response.Write "<script language=JavaScript>{alert('["&abcc&"]��������������ʹ�����֣�');parent.history.go(-1);}</script>"
	Response.End 
end if
next
yacz=idd
yinliang=abs(mg)
select case idd
	case "��"
		duboimg="<img src=../chatroom/image/da.gif>"
	case "С"
		duboimg="<img src=../chatroom/image/xiao.gif>"
	case "��"
		duboimg="<img src=../chatroom/image/dan.gif>"
	case "˫"
		duboimg="<img src=../chatroom/image/shuang.gif>"
case else
	Response.Write "<script language=JavaScript>{alert('Ѻ����Ϊ����С������˫��');parent.history.go(-1);}</script>"
	Response.End 
end select
	if yinliang<100000 or yinliang>5000000 then
		Response.Write "<script language=JavaScript>{alert('����Ѻ��10���������500������');parent.history.go(-1);}</script>"
		Response.End 
	end if
'��ʼ�ж�
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select ���� FROM �û� WHERE ����='" & un &"'"
set rs=conn.execute(sql)
if rs("����")<yinliang then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ǯ���󲻹���');parent.history.go(-1);}</script>"
	Response.End
end if
	sql= "Select count(*) As ���� from d where yl<>0 and c='��'"
	set rs=conn.execute(sql)
	durs=rs("����")
	if durs>=10 then
	        rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('["&durs&"]���ڿ��֣��Ժ�����ע��');parent.history.go(-1);}</script>"
		Response.End 		
	end if
	sql="select top 1 xm FROM d WHERE sf='ׯ��'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('����û��ׯ�ң��������߶Ĳ���������������Ͽ�����ׯ��');parent.history.go(-1);}</script>"
		Response.End 		
	end if	
	sql="select top 1 * FROM d WHERE xm='"& un &"' and c='��'"
	set rs=conn.execute(sql)
	if not(rs.eof) or not(rs.bof) then
		if rs("sf")="ׯ��" then
			temp=""&un&"��������ׯѽ����Ҫ��ʲôѽ!"
		else
			temp=""&un&"��ѹ["&rs("ydx")&"]"&rs("yl")&"�����ſ��ɣ�"
		end if
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('"&temp&"');parent.history.go(-1);}</script>"
		Response.End 
	end if
	conn.execute "update �û� set ����=����-"&yinliang&" where ����='"&un&"'"
	conn.execute "insert into d(xm,sf,yl,ydx,sj,c) values ('"&un&"','���',"&yinliang&",'"&yacz&"',now(),'��')"	
	tmprs=conn.execute("Select count(*) As ���� from d where ydx='��' and c='��'")
	dars=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from d where ydx='С' and c='��'")
	xiaors=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from d where ydx='��' and c='��'")
	danrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from d where ydx='˫' and c='��'")
	shuangrs=tmprs("����")
	tmprs=conn.execute("Select count(*) As ���� from d where yl<>0 and c='��'")
	durs=tmprs("����")


'������
if durs>=5 then
	Randomize
	m1 = Int(6 * Rnd + 1)
	Randomize
	m2 = Int(6 * Rnd + 1)
	Randomize
	m3 = Int(6 * Rnd + 1)
	sjdubo=m1+m2+m3
'����ׯ��
sql= "select * FROM d WHERE sf='ׯ��' and c='��'"
set rs=conn.execute(sql)
zhuangjia=rs("xm")
rs.close
set rs=nothing
'���Ӵ���
if m1=m2 and m2=m3 and m3=m1 then
	sql="select yl FROM d WHERE yl<>0 and sf<>'ׯ��' and c='��'"
	set rs=conn.execute(sql)
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("yl")
	rs.movenext
	loop
	rs.close
	newyingyin=int(yingyin*0.9)
	conn.execute "update �û� set ����=����+"&newyingyin&"+500000 where ����='"& zhuangjia &"'"
	conn.execute "delete * from d where  xm<>'' and c='��'"
	says="<b><font color=FF0000>�����߶Ĳ���</b></font>ׯ�ҿ���<img src=../chatroom/image/"&m1&".gif><img src=../chatroom/image/"&m2&".gif><img src=../chatroom/image/"&m3&".gif>�ƣ�"& sjdubo &"��!ׯ�ҿ������ӡ���ͨɱ��ׯ�ң�<font color=FF0000>["&zhuangjia&"]</font>���룺"&yingyin&"����ʤ�߿�10%�Ĺ�˰��"
	call chat(says)
	Response.Write "<script language=JavaScript>{alert('[�������,���ľ�]');parent.history.go(-1);}</script>"	
	Response.End 
end if

'��ʼ������
shuangyinliang=0
shuangname=""
danyinliang=0
danname=""
xiaoyinliang=0
xiaoname=""
dayinliang=0
daname=""

'����˫����
if sjdubo/2=int(sjdubo/2) then
	danshuang="<img src=../chatroom/image/shuang.gif>"
	sql="select yl,xm FROM d WHERE ydx='˫' and c='��'"
	set rs=conn.execute(sql)
	do while not rs.bof and not rs.eof
		shuangyinliang=shuangyinliang+rs("yl")
		shuangname=shuangname&rs("xm")&" "
		newyingyin=int(rs("yl")*0.9)+rs("yl")
		conn.execute "update �û� set ����=����+"&newyingyin&" where ����='"& rs("xm") &"'"
		conn.execute "update �û� set ����=����-"& rs("yl") &" where ����='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from d where ydx='˫'"
else
	danshuang="<img src=../chatroom/image/dan.gif>"
	sql="select yl,xm FROM d WHERE ydx='��' and c='��'"
	set rs=conn.execute(sql)
	do while not rs.bof and not rs.eof
		danyinliang=danyinliang+rs("yl")
		danname=danname&rs("xm")&" "
		newyingyin=int(rs("yl")*0.9)+rs("yl")
		conn.execute "update �û� set ����=����+"&newyingyin&" where ����='"& rs("xm") &"'"
		conn.execute "update �û� set ����=����-"& rs("yl") &" where ����='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from d where ydx='��'"
end if

'�Կ���С����
if sjdubo<=10 then
	daxiao="<img src=../chatroom/image/xiao.gif>"
	sql="select yl,xm FROM d WHERE ydx='С' and c='��'"
	set rs=conn.execute(sql)
	do while not rs.bof and not rs.eof
		xiaoyinliang=xiaoyinliang+rs("yl")
		xiaoname=xiaoname&rs("xm")&" "
		newyingyin=int(rs("yl")*0.9)+rs("yl")
		conn.execute "update �û� set ����=����+"&newyingyin&" where ����='"& rs("xm") &"'"
		conn.execute "update �û� set ����=����-"& rs("yl") &" where ����='"& zhuangjia &"'"
	rs.movenext	
	loop
	rs.close
	conn.execute "delete * from d where  ydx='С'"
else
	daxiao="<img src=../chatroom/image/da.gif>"
	sql="select yl,xm FROM d WHERE ydx='��' and c='��'"
	set rs=conn.execute(sql)
	do while not rs.bof and not rs.eof
		dayinliang=dayinliang+rs("yl")
		daname=daname&rs("xm")&" "
		newyingyin=int(rs("yl")*0.9)+rs("yl")
		conn.execute "update �û� set ����=����+"&newyingyin&" where ����='"& rs("xm") &"'"
		conn.execute "update �û� set ����=����-"& rs("yl") &" where ����='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from d where  ydx='��'"
end if
'��ʣ������û�����
	rs.open "select yl,xm FROM d WHERE yl<>0 and sf<>'ׯ��' and c='��'",conn
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("yl")
	rs.movenext
	loop
	rs.close
	newyingyin=int(yingyin*0.9)+500000
	conn.execute "update �û� set ����=����+"&newyingyin&" where ����='"& zhuangjia &"'"
	conn.execute "delete * from d where  xm<>'' and c='��'"
	zong=yingyin+shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	pei=shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	duboname=shuangname&danname&xiaoname&daname
	says="<b><font color=FF0000>�����߶Ĳ���</b></font>ׯ�ҿ���<img src=../chatroom/image/"&m1&".gif><img src=../chatroom/image/"&m2&".gif><img src=../chatroom/image/"&m3&".gif>�ƣ�"& sjdubo &"��!"&danshuang&daxiao&"����ע��"&zong&"����ׯ�ң�["&zhuangjia&"]���룺"&yingyin &"��,�����"&pei&"�����ϼƣ�"&yingyin-pei&"��,���У�<font color=red>"&duboname&"</font>�������Ѻ�У�ʤ�߿�10%��˰!"
    call chat(says)
    Response.Write "<script language=JavaScript>{alert('[�������,���ľ�]');parent.history.go(-1);}</script>"	
    Response.End 
    end if
    says="<b><font color=FF0000>�����߶Ĳ���</b>"&un&"</font>���Լ���С�ɰ����ó�:"&yinliang&"��������"&duboimg&"����һ���еģ�������ע���£�Ѻ��"& dars &"����ѺС:"& xiaors &"��,Ѻ����"&danrs&"��,Ѻ˫:"&shuangrs&"��������:"&(5-durs)&"�����֣�"
	call chat(says)
	Response.Write "<script language=JavaScript>{alert('[��ע�ɹ�]');parent.history.go(-1);}</script>"	
    Response.End 
	
sub chat(says)
dim newtalkarr(600) 
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr	
end sub
%>

