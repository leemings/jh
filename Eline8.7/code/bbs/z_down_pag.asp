<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="inc/chkinput.asp" -->
<!--#include file="inc/email.asp" -->
<!--#include file="inc/ubbcode.asp" -->
<%dim announceid
dim username
dim rootid
dim topic
dim mailbody
dim useremail
dim announce
dim abgcolor
dim usersign
dim sendmail
dim userclass
dim TotalUseTable
usersign=false
stats="����ʵ�"
if Cint(GroupSetting(15))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�н�����������Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	founderr=true
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ����������"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
else
	AnnounceID=request("id")
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(2,0,"","")
	if request("action")="sendmail" then
		if IsValidEmail(trim(Request.Form("mail")))=false then
			errmsg=errmsg+"<br>"+"<li>����Email�д���!</li>"
			founderr=true
		else
			useremail=trim(Request.Form("mail"))
		end if
		call announceinfo()
		if founderr then
			call dvbbs_Error()
		else
			if Forum_Setting(2)=0 then
				errmsg=errmsg+"<br>"+"<li>����̳��֧�ַ����ʼ���</li>"
				founderr=true
				call dvbbs_error()
			elseif Forum_Setting(2)=1 then
				call jmail(useremail,topic,mailbody)
				sucmsg="��ϲ�������Ĵ���ʵݷ��ͳɹ���"
				call dvbbs_suc()
			elseif Forum_Setting(2)=2 then
				call Cdonts(useremail,topic,mailbody)
				sucmsg="��ϲ�������Ĵ���ʵݷ��ͳɹ���"
				call dvbbs_suc()
			elseif Forum_Setting(2)=3 then
				call aspemail(useremail,topic,mailbody)
				sucmsg="��ϲ�������Ĵ���ʵݷ��ͳɹ���"
				call dvbbs_suc()
			end if
		end if
	else
		call pag()
	end if
	call activeonline()
end if
call footer()

sub announceinfo()
	set rs=server.createobject("adodb.recordset")
	set rs=conndown.execute("select showname from download where ID="&AnnounceID)
	if not(rs.bof and rs.eof) then
		topic=rs("showname")
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>��ָ�������������</li>"
		exit sub
	end if
	rs.close
	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: ����; FONT-SIZE: 9pt	}</style>"
	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody &"<TD valign=middle align=top>"
	mailbody=mailbody &"-&nbsp;&nbsp;<b>"&Forum_info(0)&"</b>&nbsp;&nbsp;(http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"z_down_pag.asp","")&"z_down_default.asp)<br>"
	mailbody=mailbody &"----&nbsp;&nbsp;<b>"&htmlencode(showname)&"</b>&nbsp;&nbsp;(http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"z_down_pag.asp","")&"down_list?id="&announceid&")"
	mailbody=mailbody &"</TD></TR></TBODY></TABLE><br><hr>"
	mailbody=mailbody &""&Copyright&"&nbsp;&nbsp;"&Version&""
end sub

sub pag()%>
	<html>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:60%">
	<form action="?action=sendmail&id=<%=announceid%>" method=post>
		<tr>
			<th align=left valign=middle colspan=2 height=25>����ʵ�</th>
		</tr>
		<tr>
			<td class=tablebody1 valign=middle colspan=2><b>�ѱ�������ʵݡ�</b><br>����ȷ������Ҫ�ʵݵ��ʼ���ַ��</td>
		</tr>
		<tr>
			<td class=tablebody1><b>�ʵݵ� Email ��ַ��</b></td>      
			<td class=tablebody1><input type=text size=40 name="mail"></td>
		</tr>
		<tr>
			<td colspan=2 class=tablebody2 align=center><input type=submit value="�� ��" name="Submit"></td>
		</tr>
	</form>
	</table>
	</html>
<%end sub%>
