<!--#include file="conn.asp"-->
<!--#include file="z_dgconn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="jhConst.asp"-->
<%stats="���͵��ɹ�"
call nav()
call head_var(0,0,"���������","z_dglistall.asp")

dim isfull
dim mess,says,act,towhoway,towho,addwordcolor,saycolor,addsays,saystr
isfull=false		
founderr=false

if not founduser then
	errmsg=errmsg+"<br>"+"<li>���˲����Բ��ĵ���б�"
	errmsg=errmsg+"<br>"+"<li>����<a href=login.asp><font color=blue>��¼</font></a>,��û��<a href=reg.asp><font color=blue>ע��</font></a>"
	call dvbbs_error()
else
	dim lincept,lcontent,lmedianame,lurl
	
	lincept=replace(Request.Form("incept"),"'","")
	lcontent=replace(Request.Form("content"),"'","")
	lmedianame=replace(Request.Form("medianame"),"'","")
	lurl=Request.Form("url")
	
	if lincept="" then 
		errmsg=errmsg+"<br>"+"<li>����������û���"
		founderr=true
	else
		if instr(lincept,"ȫ���Ա")>0 then
			lincept="ȫ���Ա"
			isfull=true
		else 	
			lincept=split(lincept,"|")
			if ubound(lincept)>=5 then
				errmsg=errmsg+"<br>"+"<li>���ֻ�ܷ��͸�5���û�"
				founderr=true
			end if
		end if	
	end if
	if lmedianame="" then 
		errmsg=errmsg+"<br>"+"<li>�����������"
		founderr=true
	end if
	if lurl="" or lurl="http://" then 
		errmsg=errmsg+"<br>"+"<li>����ȷ�������ֵ�ַ"
		founderr=true
	end if
	if lcontent="" then 
		errmsg=errmsg+"<br>"+"<li>������ף����"	
		founderr=true
	end if
	if len(lmedianame)>50 then '--------------�ж�����
		errmsg=errmsg+"<br>"+"<li>���������ܶ���50��"
		founderr=true	
	end if
	if len(lcontent)>200 then '--------------�ж�����
		errmsg=errmsg+"<br>"+"<li>ף���ﲻ�ܶ���200��"
		founderr=true	
	end if
	if isfull then
		if isnull(myvip) or myvip<>1 then
			if mymoney<clng(forum_user(18)) then
				errmsg=errmsg+"<br>"+"<li>��û���㹻���ֽ�Ϊȫ���Ա���"
				founderr=true
			end if
		else
			if mymoney<clng(forum_user(20)) then
				errmsg=errmsg+"<br>"+"<li>��û���㹻���ֽ�Ϊȫ���Ա���"
				founderr=true
			end if
		end if
	else
		if isnull(myvip) or myvip<>1 then
			if mymoney<clng(forum_user(19))*ubound(lincept) then
				errmsg=errmsg+"<br>"+"<li>��û���㹻���ֽ�Ϊָ���Ļ�Ա���"
				founderr=true
			end if
		else
			if mymoney<clng(forum_user(21))*ubound(lincept) then
				errmsg=errmsg+"<br>"+"<li>��û���㹻���ֽ�Ϊָ���Ļ�Ա���"
				founderr=true
			end if
		end if
	end if
	if founderr then
		call dvbbs_error()
	else
		call updata()
		call CloseDB()
		call dvbbs_suc()
	end if
end if

says="<bgsound src=wav/diangl.wav loop=1><img src=img/diangel.gif><font color=000099><b>�����ף����</b></font>"&mess
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & jhname & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
addmsg saystr
Function Yushu(a)
 Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
call activeonline()
call footer()
'===================================
sub updata()
	dim ii
	
	if isfull then
		set rs=server.createobject("adodb.recordset")
		sql="select * from media" 
		rs.open sql,connDG,3,2
		rs.addnew
		rs("sender")=membername
		rs("incept")=lincept
		rs("sendtime")=now()
		rs("content")=lcontent
		rs("medianame")=lmedianame
		rs("url")=lurl
		rs.update
		rs.close
		set rs=nothing
		if isnull(myvip) or myvip<>1 then
			conn.execute("update [user] set userWealth=userWealth-"&forum_user(18)&" where UserName='"&membername&"'")
		else
			conn.execute("update [user] set userWealth=userWealth-"&forum_user(20)&" where UserName='"&membername&"'")
		end if
                mess="<font color=#cc0000>"&jhname&"</font><font color=006600>�ڽ�����̳��[<font color=navy>ȫ���Ա</font>]����һ�׸�<img src=img/diange.gif>:<a href=../bbs/Z_dglistme.asp target=_blank title=�����ɲ鿴���><b>{"&lmedianame&"}</b></a>��ף����:<font color=ff0000><u>"&lcontent&"</u></font>��ϣ����ҵ��<a href=../bbs/Z_dglistme.asp target=_blank title=�����ɲ鿴���><b>����</b></a>�鿴��裡</font>"
		sucmsg="<li>���ף���ѳɹ�����!<br><li>������̳���еĻ�Ա�����˵��ף��"	
	else	
		for ii=0 to ubound(lincept)
			set rs=server.createobject("adodb.recordset")
			sql="select username from [user] where username='"&lincept(ii)&"'"
			rs.open sql,conn,1,1
			if rs.eof and rs.bof then
				sucmsg=sucmsg+"<br>"+"<li>��̳û��<font color=red>["&lincept(ii)&"]</font>����û�"
			else
				rs.close
				sql="select * from media" 
				rs.open sql,connDG,3,2
				rs.addnew
				rs("sender")=membername
				rs("incept")=lincept(ii)
				rs("sendtime")=now()
				rs("content")=lcontent
				rs("medianame")=lmedianame
				rs("url")=lurl
				rs.update
				rs.close
				set rs=nothing
				dim sender,title,body
				sender=membername
				title="�͸�����ף��"
				body="[color=green]"&[membername]&"[/color] ����һ�׸� [color=navy]"&lmedianame&"[/color] ���㣡"&chr(10)&"ף���[color=blue]"&lcontent&"[/color] "&chr(10)&"[B][URL=Z_dglistme.asp]�������鿴���[/URL][/B]"
				sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&lincept(ii)&"','"&sender&"','"&title&"','"&body&"',Now(),0,1)"
				conn.Execute(sql)
				if isnull(myvip) or myvip<>1 then
					conn.execute("update [user] set userWealth=userWealth-"&forum_user(19)&" where UserName='"&membername&"'")
				else
					conn.execute("update [user] set userWealth=userWealth-"&forum_user(21)&" where UserName='"&membername&"'")
				end if
                mess="<font color=#cc0000>"&jhname&"</font><font color=006600>�ڽ�����̳������[<font color=navy>"&lincept(ii)&"</font>]����һ�׸�<img src=img/diange.gif>:<a href=../bbs/Z_dglistme.asp target=_blank title=�����ɲ鿴���><b>{"&lmedianame&"}</b></a>��ף����:<font color=ff0000><u>"&lcontent&"</u></font>��ϣ������[<font color=navy>"&lincept(ii)&"</font>]���<a href=../bbs/Z_dglistme.asp target=_blank title=�����ɲ鿴���><b>����</b></a>�鿴��裡</font>"
				sucmsg=sucmsg+"<li>���ף���ѳɹ�����!ϵͳ����<font color=green>["&lincept(ii)&"]</font>��������Ϣ֪ͨ"
			end if	
		next
	end if		
end sub
%>