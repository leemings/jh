<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_Visual_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="z_plus_check.asp"-->
<!--#include file="jhconst.asp"-->
<%
'=========================================================
' Visual photo For ˮ������������ϵͳ2.0
' File: z_Visual_photo.asp
' Version:1.0b
' Date: 2003-5-10
' Script Written by TonyHe
'=========================================================
' dhjh.Net
' Web: http://www.dhjh.net,http://www.dhjh.net/club
' Email: hegq@v889.com
' MSN: hegq@msn.com
' QQ:611496
' ��Ȩû�У���ӭ�޸��Ż������ṩ�Ż����ļ�����^_^
'=========================================================
%>
<%
dim pid,curpage,act,app,uid,pas,pmyname,lst,content
dim prs,PageCount,curcount
if not founduser or membername="" then
	errmsg=errmsg+"<br>"+"<li>����û��<a href=login.asp>��¼����</a>������ʹ�ø��������Ӱ���ܡ��������û��<a href=reg.asp>ע��</a>������<a href=reg.asp>ע��</a>��"
	founderr=true
end if

act=checkstr(request("act")) '�ύ������0Ϊ�����Ӱ1Ϊͬ���Ӱ2�ۿ���Ƭ
app=checkstr(request("app")) '�ύ������0Ϊ�����Ӱ1Ϊͬ���Ӱ
lst=checkstr(request("lst")) 'Ӱ�������
pas=checkstr(request("pas")) '����
pid=checkstr(request("pid")) '��Ӱ��id
uid=checkstr(request("uid")) '�û�id
curPage=request("page") 'ҳ��

if not isInteger(pid) then pid=""
if not isInteger(lst) then lst=""
if not isInteger(uid) then uid=""
if act="" or not isInteger(act) then act=3
if not isInteger(app) then app=""

if uid<>"" then uid=clng(uid)
if act<>"" then act=clng(act)
if app<>"" then app=clng(app)
if pid<>"" then pid=clng(pid)
if lst<>"" then lst=clng(lst)

if curpage="" or not isInteger(curpage) then
	curpage=1
else
	curpage=clng(curpage)
end if
if uid="" then
	uid=userid
	pmyname=membername
	else
	pmyname=getmyname(uid)
end if
stats="�������"
if founderr=true then
	call nav()
        call head_var(2,0,"","")
	call dvbbs_error()
else
                        select case act
                        case 6 'һ����Ƭ
                                call nav()
                                stats="�����Ƭ"
                                call photomain()
			        call photoone(pid,pas)
                        case 5 '�Ƽ�ɾ��
                                call nav()
                                stats="��Ƭ����"
                                call photomain()
			        call photodel(pid,lst)
                        case 4 'vote��ƬͶƱ
                                call nav()
                                stats="����ͶƱ"
                                call photomain()
			        call photovote()
                        case 3 'top20�����������
                                call nav()
                                stats="�������"
                                call photomain()
			        call phototop()
			case 2 '�ۿ�����Ҫ���ú�Ӱid������
			        stats="�������"
                                call photomain()
			        call photosee(pid,pas)
			case 1 'ͬ��ܾ�
			        call nav()
                                stats="��Ӱ����"
                                call photomain()
				call photoapp(uid,act,app,pid)
			case 0 '�������
			        call nav()
                                stats="�����Ӱ"
                                call photomain()
			        if app="3"  then
			        call photosubgo(uid,app,lst)
			        elseif app="2" then
			        call photosubok(uid,act,app,pid)
			        else
			        call photosub()
			        end if
			case else '�ҵ����
			        call nav()
                                stats="�ҵ����"
                                call photomain()
			        call photolist(lst)
			end select
response.write "                       </td>" & vbCrLf
response.write "			</tr>" & vbCrLf
response.write "			</table>"
end if
call footer()


sub photomain()
call head_var(0,0,"�������","z_Visual_photo2.asp?act=9")
response.write "<table cellpadding=2 cellspacing=1 align=center class=tableborder1>" & vbCrLf
response.write "<tr>" & vbCrLf
response.write "	<td align=center valign=top class=tablebody2 width=160>"
response.write "                        <table border=0 cellspacing=1 cellpadding=0 width=140 bgcolor="&forum_body(27)&">"			
response.write "                        <tr height=25 valign=middle>" & vbCrLf
response.write "			<td class=tablebody1>&nbsp;<a href=z_Visual_photo2.asp?act=9&uid="&uid&"&lst=>"&pmyname&"�����</a></td>" & vbCrLf
response.write "			</tr>" & vbCrLf
response.write "			<tr height=25 valign=middle>" & vbCrLf
response.write "			<td class=tablebody1>&nbsp;<a href=z_Visual_photo2.asp?act=9&uid="&uid&"&lst=1>д�漯</a></td>" & vbCrLf
response.write "			</tr>" & vbCrLf
response.write "			<tr height=25 valign=middle>" & vbCrLf
response.write "			<td class=tablebody1>&nbsp;<a href=z_Visual_photo2.asp?act=9&uid="&uid&"&lst=0>��ͥ��</a></td>" & vbCrLf
response.write "			</tr>" & vbCrLf
response.write "			<tr height=25 valign=middle>" & vbCrLf
response.write "			<td class=tablebody1>&nbsp;<a href=z_Visual_photo2.asp?act=9&uid="&uid&"&lst=3>���Ѽ�</a></td>" & vbCrLf
response.write "			</tr>" & vbCrLf
response.write "			<tr height=25 valign=middle>" & vbCrLf
response.write "			<td class=tablebody1>&nbsp;<a href=z_Visual_photo2.asp?act=0&app=3&lst=2 title=����˴�����д��>��Ҫд��</a></td>" & vbCrLf
response.write "			</tr>" & vbCrLf
response.write "			<tr height=25 valign=middle>" & vbCrLf
response.write "			<td class=tablebody1>&nbsp;<a href=z_Visual_photo2.asp?act=0>��������</a></td>" & vbCrLf
response.write "			</tr>" & vbCrLf
response.write "			<tr height=25 valign=middle>" & vbCrLf
response.write "			<td class=tablebody1>&nbsp;<a href=z_Visual_photo2.asp?act=3>�������</a></td>" & vbCrLf
response.write "			</tr>" & vbCrLf
response.write "			</table>" & vbCrLf
response.write "			</td>" & vbCrLf
response.write "			<td align=center valign=top class=tablebody1 width=""95%"">"
end sub

sub photoone(vpid,vpas)'�ҵ���Ƭ
dim myname
dim mypid,mypas
if lst="" then lst="0"
if uid=""  then 
	uid=userid
        myname="��"
        else
myname=getmyname(uid)
end if
%>
<a name=visualtop>
<table border=0 cellspacing=1 cellpadding=0 bgcolor=<%=Forum_Body(27)%> width=98%>
<tr height=25 valign=middle>
<th valign=middle>
<a href="z_Visual_photo2.asp?act=9&uid=<%=uid%>&lst=3"><%=myname%>�ĺ�Ӱ��</a> -  
<a href="z_Visual_photo2.asp?act=9&uid=<%=uid%>&lst=0"><%=myname%>�ļ�ͥ��</a> - 
<a href="z_Visual_photo2.asp?act=9&uid=<%=uid%>&lst=1"><%=myname%>��д����</a> -  
</th></tr>
<tr height=25 valign=middle>
<td valign=middle width=100%>����<%=myname%>�Ƽ�����Ƭ 
<%
                set prs=server.createobject("ADODB.Recordset")
		sql="select id,book,date,cou from [Visual_photo2] where pok=1 and id="&vpid
		set prs=conn.execute(sql)
		if prs.eof and prs.bof then
		response.write "<br>��ʱ��û���κ��û�����</td>"
                else
                mypid=prs("id")
		%>
		</td></tr>
		<tr><td class=tablebody1 width=45%>
					<table width=100% border=0 cellspacing=1 cellpadding=0>
					<tr><td align=center>
					<%
					call photosee(mypid,vpas)
					%>					
                                        </td></tr>
                                        <tr><td class=tablebody2>
                                        <%
                                        response.write "&nbsp;&nbsp;<a href=""z_Visual_photo2.asp?act=4&lst=1&pid="&mypid&"&uid=""&userid&"" title=""����������󣬸���һ���ʻ�����������"&GroupSetting(47)&"���Ǯ""><img src="&Forum_info(7)&Forum_BoardPic(6)&" border=0>�ʻ�</a>&nbsp;&nbsp;<a href=""z_Visual_photo2.asp?act=4&lst=0&pid="&mypid&"&uid=""&userid&"" title=""����������󣬸���һ����������������"&GroupSetting(47)&"���Ǯ""><img src="&Forum_info(7)&Forum_BoardPic(7)&" border=0>����</a>(����<font color="&Forum_body(8)&">"&prs("cou")&"</font>)"
                                        %>
                                        </td></tr>
                                        <tr><td class=tablebody2>&nbsp;&nbsp;��Ӱ���ڣ�
                                        <%=prs("date")%><br>&nbsp;&nbsp;
                                        <%=prs("book")%>
                                        </td></tr>
                                        </table>
                </td></tr></table><%
			
end if
prs.close
set prs=nothing
end sub

sub photodel(vpid,vlst)'��Ƭ����
dim phtname,myid1,myid2
select case vlst
case "1" '�Ƽ�
	phtname=checkstr(request("phtname"))
	if phtname="" or getmyid(phtname)="" then
   	   response.write"<br>"+"<li>�Ƽ����Է���������Ҫд��ȷ���ְ�^_^"
	   exit sub
	end if
	content=phtname&"�����ã�"&chr(10)
        content=content & "��������[color=blue]"&membername&"[/color][b]�Ƽ�[/b][color=red]������һ�����գ�[/color]"&chr(10)	
	content=content & "[align=center][URL=z_visual_photo2.asp?act=6&pas="&pas&"&pid="&vpid&"&uid="&uid&"][color=blue][b][u]��������ǰȥ�ۿ�[/u][/b][/color][/URL][/align]"
        call messto(phtname,"�������","������������Ŷ",content)
        response.write"<br>"+"<li>�Ѿ�����֪ͨ�Է������^_^"
        call photosee(pid,pas)
case "0" 'ɾ��
        if uid="" or uid<>userid or getmyname(uid)="" then
	response.write"<br>"+"<li>���ִ�������µ�½^_^"
	exit sub
	end if
	
        set prs=server.createobject("adodb.recordset")
        sql="select myid from [Visual_photo2] where id="&vpid
        set prs=conn.execute(sql)
        if prs.eof or prs.bof then
        response.write"<Br>"+"<li>δ�ҵ���Ҫ�������Ƭ��"
        prs.close
        set prs=nothing
   	exit sub
	else
	myid1=prs("myid")
	myid2=split(myid1,"|")
	end if
        
        if myid2(0)=cstr(userid) or master then
        conn.execute("Delete FROM Visual_photo2 WHERE id="&vpid&"")
        response.write"<Br>"+"<li>�ɹ�ɾ����Ƭ��"
        else
        response.write"<Br>"+"<li>��������Ƭ��ӵ���ߣ�����ɾ����Ƭ��"
   	exit sub
        end if
	end select
set prs=nothing
end sub

sub photovote()'��ƬͶƱ
dim mycou,myid1,myid,getmoney,mylb
set prs=server.createobject("ADODB.Recordset")
sql="select * from [Visual_photo_v] where uid="&userid&" and pid="&pid
set prs=conn.execute(sql)
if not(prs.eof or prs.bof) then
response.write "<br>"+"<li>����Ƭ���Ѿ�ͶƱ��"
	exit sub
end if	
sql="select myid,cou,lb from [Visual_photo2] where pok=1 and vopen=0 and id="&pid
set prs=conn.execute(sql)
if prs.eof or prs.bof then
	response.write "<br>"+"<li>��ʱ���ܸ�����ƬͶƱ�����ߺ�Ӱ��δ��ɡ�"
	prs.close
        set prs=nothing
	exit sub
	else
	mycou=prs("cou")
	myid=prs("myid")
	myid1=split(myid,"|")
	mylb=prs("lb")
	
	if myid1(0)=cstr(userid) or myid1(1)=cstr(userid) or myid1(2)=cstr(userid) then
        founderr=true
        end if
        
	if founderr=true then
	response.write "<br>"+"<li>�����ܸ��Լ�����Ƭ���ӰͶƱ��"
	exit sub
        end if

	if jhdj<15 then
	response.write "<br>"+"<li>���Ľ����ȼ�����15�������ܽ���ͶƱ��"
	exit sub
        end if
        
	getmoney=Cint(GroupSetting(47))
	if mymoney<getmoney then
	response.write "<br>"+"<li>��û���㹻��Ǯ���������"
	exit sub
	end if
	if lst<>0 then'�ʻ�
	mycou=mycou+1
	else
	mycou=mycou-1
	end if
	
end if
prs.close
set prs=nothing
conn.execute("update [Visual_photo2] set cou='"&mycou&"' WHERE id="&pid&"")
conn.execute("update [user] set userWealth=userWealth-"&getmoney&" where userid="&userid)
conn.execute("insert into Visual_photo_v(uid,pid) values("&userid&","&pid&")")
response.write "��ϲ��ͶƱ�ɹ�!<br><br>" & vbCrLf
call photosee(pid,"")
end sub

sub phototop()'top20�����������
dim totalrec,mypid,srs
%>
<a name=visualtop>
<table border=0 cellspacing=1 cellpadding=0 bgcolor=<%=Forum_Body(27)%> width=100%>
<tr height=25 valign=middle>
<th valign=middle>
<a href="z_Visual_photo2.asp?act=3&lst=4">[����������TOP20����]</a>   
</th></tr>
<tr height=25 valign=middle>
<td valign=middle class=tablebody2 align=center>
<a href="z_Visual_photo2.asp?act=3&lst=3">������</a> -  
<a href="z_Visual_photo2.asp?act=3&lst=2">˫����</a> - 
<a href="z_Visual_photo2.asp?act=3&lst=1">д����</a> -  
<a href="z_Visual_photo2.asp?act=3&lst=0">������</a> - 
<a href="z_Visual_photo2.asp?act=3&lst=6">ȫ����</a>
</td></tr>
<tr height=25 valign=middle>
<td valign=middle width=100%> 
<%
set srs=server.createobject("ADODB.Recordset")
                        select case lst
                        case "4" '������TOP20����
                        sql="select top 20 id,book,date,cou from [Visual_photo2] where pok=1 and vopen=0 order by cou desc"
                        %>  ���������������<%
                        'sql="select id,book,date,cou from [Visual_photo2] where pok=1 and vopen=0 order by id desc"
                        case "3" 'top20�����������
                        %>  TOP20����������<%
                        sql="select top 20 id,book,date,cou from [Visual_photo2] where pok=1 and vopen=0 and lb=3 order by cou desc"			        
			case "2" 'top20�����������
			%>  TOP20����˫����<%
                        sql="select top 20 id,book,date,cou from [Visual_photo2] where pok=1 and vopen=0 and lb=2 order by cou desc"			        
			case "1" 'top20�����������
			%>  TOP20����д����<%
                        sql="select top 20 id,book,date,cou from [Visual_photo2] where pok=1 and vopen=0 and lb=1 order by cou desc"				
			case "0" 'top20�����������
			%>  TOP20����˫����<%
                        sql="select top 20 id,book,date,cou from [Visual_photo2] where pok=1 and vopen=0 and lb=0 order by cou desc"			        
			case else 'ȫ����Ƭ�б�
			%>  �������ȫ����Ƭ�б�<%
                        sql="select id,book,date,cou from [Visual_photo2] where pok=1 and vopen=0 order by id desc"
			end select
                        srs.open sql,conn,1,1
                        if srs.eof and srs.bof then
		               response.write "<br>��ʱ��û���κ��û�����</td>"
                        else
                        %></td></tr><tr><td>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr valign=middle>
			<% 
			                totalrec=srs.recordcount
                                        if totalrec mod 10=0 then
						pagecount=totalrec \ 10
					else
						pagecount=totalrec \ 10+1
					end if
					curcount=0
					dim iic
					iic=0
					if curpage > pagecount then curpage = pagecount
						if curpage < 1 then curpage=1
						srs.PageSize=10
						srs.AbsolutePage=CurPage
					do while not srs.eof and curcount<10
					mypid=srs("id")
					%>
					<td class=tablebody1 width=45%>
					<table width=100% border=0 cellspacing=1 cellpadding=0><tr><td align=center>
					<%call photosee(mypid,"")%>					
                                        </td></tr>
                                        <tr><td class=tablebody2>
                                        <%
                                        response.write "&nbsp;&nbsp;<a href=""z_Visual_photo2.asp?act=4&lst=1&pid="&mypid&"&uid="&userid&""" title=""����������󣬸���һ���ʻ�����������"&GroupSetting(47)&"���Ǯ""><img src="&Forum_info(7)&Forum_BoardPic(6)&" border=0>�ʻ�</a>&nbsp;&nbsp;<a href=""z_Visual_photo2.asp?act=4&lst=0&pid="&mypid&"&uid="&userid&""" title=""����������󣬸���һ����������������"&GroupSetting(47)&"���Ǯ""><img src="&Forum_info(7)&Forum_BoardPic(7)&" border=0>����</a>(������<font color="&Forum_body(8)&">"&srs("cou")&"</font>)"
                                        if master then
                                        response.write "&nbsp;&nbsp;<a href=""z_Visual_photo2.asp?act=5&lst=0&pid="&mypid&"&uid=""&userid&"" title=""ɾ���⸱��Ӱ"">ɾ��</a>"
                                        end if
                                        %>
                                        </td></tr>
                                        <tr><td class=tablebody2>&nbsp;&nbsp;��Ӱ���ڣ�
                                        <%=srs("date")%><br>&nbsp;&nbsp;
                                        <%=srs("book")%>
                                        </td></tr>
                                        <form action="z_Visual_photo2.asp?act=5&pid=<%=mypid%>&lst=1&uid=<%=userid%>&pas=" name="subsee" method=post>
                                        <tr><td class=tablebody2>&nbsp;&nbsp;
                                        <INPUT name=phtname type=text> &nbsp;<input type=submit name=submit value="�Ƽ�������">
                                        </td></tr>
                                        </form>
                                        </table></td>
                                        <%
                                        curcount=curcount+1
                                        iic=iic+1
                                        if iic mod 2=0 then
                                        iic=0%>
                                        <tr valign=middle>
                                        <%end if
					srs.movenext
					loop
					
			%>
			</tr></table></td>
			<%		
			end if
			%>
			</tr><tr><td width=100%>
			<table width=100% border=0 cellspacing=0 cellpadding=0>
					<tr height=25 valign=middle>
						<td valign=middle class=tablebody1>&nbsp;&nbsp;ҳ�Σ�<b><%=curPage%></b> / <b><%=pagecount%></b> ҳ ÿҳ��<b>10</b> ���ƣ�<b><%=totalrec%></b></td>
						<td valign=middle class=tablebody1><div align=right>��ҳ��<%
							call DispPageNum(curpage,pagecount,"?act=3&lst="&lst&"&page=","#VisualTop")
						%>&nbsp;&nbsp;</td>
					</tr>
			</table>
			</td></tr></table>
			<%
srs.close
set srs=nothing
end sub

sub photosee(pid,pas) '�ۿ�����Ҫ���ú�Ӱid������
dim myid,myv,myc,mybake,myopen,mypass,mylb,mybook,mydate
dim myid_n,myv_n,myc_n,pwidth,myname1,myname2,myname3,mysex1,mysex2,mysex3
dim trs
set trs=server.createobject("ADODB.Recordset")
sql="select * from [Visual_photo2] where pok=1 and id="&pid
set trs=conn.execute(sql)
if trs.eof or trs.bof then
	response.write"<br>"+"<li>��ʱû�����뿴����Ƭ�����ߺ�Ӱ��δ��ɡ�"
	trs.close
        set trs=nothing
	exit sub
else
        myid=trs("myid")
        myv=trs("myv")
        myc=trs("myc")
        mybake=trs("bake")
        myopen=trs("vopen")
        mypass=trs("pass")
        mylb=trs("lb")
        mybook=trs("book")
        mydate=trs("date")
        if mybake="" or isnull(mybake) then
	pwidth=140
	mybake=""
	else
	pwidth=280
        end if
end if
trs.close        
set trs=nothing 
        myid_n=split(myid,"|")
        myv_n=split(myv,"��")
        myc_n=split(myc,"|")
        dim passok
        
        if myopen=0 then '��Ƭ��Ҫ����
        if myid_n(0)=cstr(userid) or myid_n(1)=cstr(userid) or myid_n(1)=cstr(userid) or master then '��Ƭ�����˿��Բ���Ҫ����
	passok=true
	else
	if pas<>mypass then
		passok=false
		else
		passok=true
	end if
        end if
        end if
        
        if passok=false then '��ᱣ����Ҫ��ȷ������
	response.write"<br>"+"<li>��Ƭ����</a>����ʱ���ܹۿ���"
        exit sub
        end if

        myname1=getmyname(myid_n(0))
	myname2=getmyname(myid_n(1))
	myname3=getmyname(myid_n(2))
	mysex1=getmysex(myid_n(0))
	mysex2=getmysex(myid_n(1))
	mysex3=getmysex(myid_n(2))
	call photolook(mylb,pwidth,mybake,myname1,myname2,myname3,mysex1,mysex2,mysex3,myv_n(0),myv_n(1),myv_n(2),myc_n(0),myc_n(1),myc_n(2),mydate)		

end sub

sub photoapp(uid,act,app,pid)'ͬ��ܾ�
dim myid1,myapp,myid2,mylb
if app<>"" and pid<>"" and uid<>"" then'ͬ����߾ܾ���Ӱ
	if uid="" or uid<>userid or getmyname(uid)="" then
	response.write"<br>"+"<li>���ִ�������µ�½^_^"
	exit sub
	end if
        set prs=server.createobject("adodb.recordset")
        sql="select myid,app,lb from [Visual_photo2] where pok=0 and id="&pid
        set prs=conn.execute(sql)
        if prs.eof or prs.bof then
        response.write"<Br>"+"<li>�������Ӱ�Ѿ����˾ܾ������Ѿ���ɡ�"
   	prs.close        
        set prs=nothing 
        exit sub
	else
	myid1=prs("myid")
	myapp=prs("app")
	myid2=split(myid1,"|")
	mylb=prs("lb")
	if myapp=mylb then
	response.write"<Br>"+"<li>���Ѿ�ͬ���Ӱ�ˣ���Ҫ�ٵ����"
   	prs.close        
        set prs=nothing 
        exit sub
	end if
        end if
        dim appok
                        select case mylb
                        case 3 '
                        if myid2(1)=cstr(userid) or myid2(2)=cstr(userid) then
                        appok=true
                        end if
			case 2,0 '
			if myid2(1)=cstr(userid) then
                        appok=true
                        end if
			end select
        
        if appok=true then
        
        if app<>1 then'�ܾ�
        conn.execute("Delete FROM Visual_photo2 WHERE id="&pid&"")
        content=getmyname(myid2(0))&"�����ã�"&chr(10)
        content=content & "��������[color=blue]"&membername&"[/color][b]�ܾ�[/b]������Ӱ[color=red]���κ�Ӱȡ��[/color]"	
        call messto(getmyname(myid2(0)),"�������","�ܾ���Ӱ",content)
        else '����
        conn.execute("update Visual_photo2 set app=app+1 WHERE id="&pid&"")
	if mylb=myapp+1 then
	content=getmyname(myid2(0))&"�����ã�"&chr(10)
        content=content & "��������[color=blue]"&membername&"[/color][b]ͬ��[/b]������Ӱ[color=red]���κ�Ӱ�����Ѿ�ȫ��ͬ�⣻[/color]"&chr(10)	
	content=content & "[align=center][URL=z_visual_photo2.asp?act=0&app=2&uid="&myid2(0)&"&pid="&pid&"][color=blue][b][u]����������ɺ�Ӱ[/u][/b][/color][/URL][/align]"
        call messto(getmyname(myid2(0)),"�������","����ɺ�Ӱ",content)
        else
        content=getmyname(myid2(0))&"�����ã�"&chr(10)
        content=content & "��������[color=blue]"&membername&"�Ѿ�[/color][b]ͬ��[/b]������Ӱ��"	
        call messto(getmyname(myid2(0)),"�������","ͬ���Ӱ",content)
        end if
        end if
        
        else	
                response.write"<Br>"+"<li>�����Ǳ�����ģ���Ȩ������"
   	        exit sub
        end if
        response.write "��ϲ�������ɹ�" & vbCrLf
end if
prs.close
set prs=nothing
end sub

sub photosub()'����
%>
<br>
<table cellpadding=2 cellspacing=0 align=center width=90%>
<tr align=center>
<th width="100%" height=25 colspan=2>�����Ӱ</th>
</tr>
<tr>
<td width="100%" class=tablebody1 colspan=2>
<b>������</b><br><br>
<li>��ȷ��д��Ҫ��Ӱ�ĶԷ����֡�
<li>���˺�Ӱ����ѡ�񱳾���
<li>˫���Լ����޺�Ӱ��ѡ�񱳾�����ѡ����Ƭ�������˵ı���Ϊ׼��
</td></tr>

<form action="z_Visual_photo2.asp?act=0&app=3&lst=1&uid=<%=userid%>" name="subpht" method=post onsubmit='return(logi());'>
    <tr>
    <th valign=middle colspan=2 align=center height=25>�����������Ӱ����������</td></tr>
    <tr>
    <td valign=middle class=tablebody1>�����������������</td>
    <td valign=middle class=tablebody1><INPUT name=phtname type=text> &nbsp; ��|�������TonyHe|����|Test</td></tr>
    <td class=tablebody1 valign=top width=30% ><b>��Ӱѡ��</b><BR> ��ѡ���Ӱ���͡�</td>
    <td valign=middle class=tablebody1>
                <input type=radio name=phtlb value=0 checked>���޼���<br>
                <input type=radio name=phtlb value=2>˫�˺�Ӱ<br>
                <input type=radio name=phtlb value=3>���˺�Ӱ<br>                </td></tr>
    <tr>
    <td valign=top width=30% class=tablebody1><b>�Ƿ���</b><BR> ������ѡ���ܣ����ѡ���ܣ����Ӱ��Ҫ��������߸�֪�Է�����ſ��������</td>
    <td valign=middle class=tablebody1>                
                <input type=radio name=phtbm value=2 checked>�������<br>
                <input type=radio name=phtbm value=1>˽�����<br>
                </td></tr>
    <tr>
    <td valign=middle class=tablebody1>��������������</font></td>
    <td valign=middle class=tablebody1><INPUT name=phtpass type=password> &nbsp; �����ѡ����˽����أ�����������</td></tr>
    <tr>
    <td valign=middle class=tablebody1>��������������</font></td>
    <td valign=middle class=tablebody1><TEXTAREA name=phtContent cols=80 rows=3 wrap=VIRTUAL title=�������Ӱ���ɻ���˵��></TEXTAREA> &nbsp; </td></tr>
    <tr>
    <tr>
    <td valign=middle class=tablebody1>ѡ�񱳾�ͼ��<br>
    <script language="javascript">
<!--
function showimage()
{
	document.images.useravatars.src = document.subpht.phtbake.options[document.subpht.phtbake.selectedIndex].value;
}

function logi()
{
if(subpht.phtname.value==""){alert("�û�������Ϊ��!");return false;}
if(subpht.phtContent.value==""){alert("������д�������ɻ���˵��!");return false;}
if(subpht.phtbm.checked){
if(subpht.phtpass.value==""){alert("���벻��Ϊ��!");return false;}
}
if(subpht.phtlb.checked){
if(subpht.phtbake.value==""){alert("���˺�Ӱ����ѡ�񱳾�!");return false;}
}
}
//-->
</script>
<select name="phtbake" size="1" onChange="showimage()">
<option value="" selected>����Ҫ����</option>
<option value="images/img_visual/show/1757.gif">1757</option>
<option value="images/img_visual/show/1758.gif">1758</option>
<option value="images/img_visual/show/1759.gif">1759</option>
<option value="images/img_visual/show/1762.gif">1762</option>
<option value="images/img_visual/show/1783.gif">1783</option>
<option value="images/img_visual/show/1784.gif">1784</option>
<option value="images/img_visual/show/1785.gif">1785</option>
<option value="images/img_visual/show/1787.gif">1787</option>
</select>
</font></td>
    <td valign=middle class=tablebody1>
    <IMG src="" border=0 width=52 height=0><img src="images/img_visual/show/blank.gif" align="absmiddle" name="useravatars" border=0>
    </td></tr>
    <tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name=submit value="���ͺ�Ӱ����"></td></tr>
</form>
</table>
<%
end sub

sub photosubgo(uid,app,lst)'����
dim phtname,phtlb,phtbm,phtpass,phtContent,phtbake,phtname_n
dim getid1,getid2,myname1,myname2
     if uid="" or getmyname(uid)="" or userid<>uid then
	response.write"<br>"+"<li>���ִ������������^_^"
	exit sub
     end if
     if getmyv(uid)="" then
	response.write"<Br>"+"<li>����û���������������ô����˼����������˺�Ӱ^_^"
   	exit sub
     end if
select case lst
case 1 '�������
        phtname=checkstr(request("phtname")) 
   	phtlb=checkstr(request("phtlb")) 
   	phtbm=checkstr(request("phtbm")) 
   	phtpass=checkstr(request("phtpass")) 
   	phtContent=checkstr(request("phtContent")) 
   	phtbake=checkstr(request("phtbake"))
   	if isnull(phtlb) or not isInteger(phtlb) then phtlb=""
        if isnull(phtbm) or not isInteger(phtbm) then phtbm=""
        if chkpost=false or phtlb="" or phtbm="" then
	errmsg=errmsg+"<br>"+"<li>�����ύ�����뷵�ش����ύ��"
	founderr=true
        end if
   	                if phtname="" then
   	                        response.write"<br>"+"<li>�����Ӱ�û�������Ϊ��^_^"
	                        exit sub
	                end if
	                
   	                if phtbm=1 and phtpass="" then
	                	response.write"<br>"+"<li>ѡ���ܣ��������д����^_^"
	                        exit sub
	                end if
	                
	                select case phtlb
                        case 3 '
                        if phtbake="" then
	                	response.write"<br>"+"<li>���˺�Ӱ����ѡ�񱳾�^_^"
	                        exit sub
	                end if 
                        phtname_n=split(phtname,"|")
   	                getid1=getmyid(phtname_n(0))
   	                getid2=getmyid(phtname_n(1))
   	                myname1=phtname_n(0)
   	                myname2=phtname_n(1)
	                	if myname1="" or myname2="" or getid1="" or getid2="" then
	                	response.write"<br>"+"<li>���˺�Ӱ����ȷ��дҪ��������֣�������|���^_^"
	                        exit sub
	                        end if
	                        if getmyv(getid1)="" or getmyv(getid2)="" then
	                        response.write"<Br>"+"<li>����������ѻ�û���������������ô����˼�����Ӱ^_^"
   	                        exit sub
	                        end if
	                         if getid1=userid or getid2=userid then
	                        response.write"<Br>"+"<li>���Լ���Ӱ������������^_^"
   	                        exit sub
	                        end if	
	                        
			case 0,2 '
			myname1=phtname
   	                myname2=""
			if phtlb=0 and not getmyw(phtname,userid) then
	               	response.write"<Br>"+"<li>�������װ��ģ���ô���������ͥ���޺�Ӱ��^_^"
   	                exit sub
	                end if
	                getid1=getmyid(phtname)
	                getid2=""
	                if getid1="" then
	                response.write"<Br>"+"<li>����ȷ��дҪ���������"
   	                exit sub
	                end if	
	                        
	                if getmyv(getid1)="" then
	                response.write"<Br>"+"<li>����������ѻ�û������������ô����˼�����Ӱ^_^"
   	                exit sub
	                end if
	                  if getid1=userid  then
	                        response.write"<Br>"+"<li>���Լ���Ӱ������������^_^"
   	                        exit sub
	                        end if	       
			end select
			
	                if phtContent="" or strLength(phtContent)>50 or strLength(phtContent)<8 then
   	                        response.write"<br>"+"<li>�����Ӱ�����ɲ��ܶ���50���ֻ��߲�������8����^_^"
	                        exit sub
	                end if
	                if chkpost=false then
	                errmsg=errmsg+"<br>"+"<li>�����ύ�����뷵�ش����ύ��"
	                founderr=true
                        end if
set prs=server.createobject("ADODB.Recordset")
sql="select * from [Visual_photo2]"	
prs.open sql,conn,1,3
        prs.addnew
        prs("myid")=uid&"|"&getid1&"|"&getid2
	prs("myv")=getmyv(uid)&"��"&getmyv(getid1)&"��"&getmyv(getid2)
	prs("myc")=getmyc(uid)&"|"&getmyc(getid1)&"|"&getmyc(getid2)
	prs("bake")=phtbake
	if phtbm<>1 then
		phtbm=0
		prs("pass")=""
		else
		phtbm=1
		prs("pass")=md5(phtpass)
	end if
	prs("vopen")=phtbm
	
	prs("lb")=phtlb
	if phtlb=0 then
	prs("app")=-1
		else
	prs("app")=1
        end if
	prs("book")=phtContent
	prs.update
prs.close
set prs=nothing
set prs=conn.execute("select top 1 id from [Visual_photo2] order by id desc")
pid=prs(0)

if phtlb<>3 then
content=myname1&"�����ã�"&chr(10)
content=content & "��������[color=blue]"&membername&"[/color][b]����[/b]������Ӱ[color=red]���������ǣ�[/color]"&chr(10)
content=content & phtContent &chr(10)
content=content & "[align=center][URL=z_visual_photo2.asp?act=1&app=1&uid="&getid1&"&pid="&pid&"][color=blue][b][u]�������ͬ���Ӱ[/u][/b][/color][/URL][/align]"&chr(10)
content=content & "[align=center][URL=z_visual_photo2.asp?act=1&app=0&uid="&getid1&"&pid="&pid&"][color=red][b][u]�������ܾ���Ӱ[/u][/b][/color][/URL][/align]"
call messto(myname1,"�������","�����Ӱ",content)
else
content=myname1&"�����ã�"&chr(10)
content=content & "��������[color=blue]"&membername&"[/color][b]����[/b]������Ӱ[color=red]���������ǣ�[/color]"&chr(10)
content=content & phtContent &chr(10)
content=content & "[align=center][URL=z_visual_photo2.asp?act=1&app=1&uid="&getid1&"&pid="&pid&"][color=blue][b][u]�������ͬ���Ӱ[/u][/b][/color][/URL][/align]"&chr(10)
content=content & "[align=center][URL=z_visual_photo2.asp?act=1&app=0&uid="&getid1&"&pid="&pid&"][color=red][b][u]�������ܾ���Ӱ[/u][/b][/color][/URL][/align]"
call messto(myname1,"�������","�����Ӱ",content)

content=myname2&"�����ã�"&chr(10)
content=content & "��������[color=blue]"&membername&"[/color][b]����[/b]������Ӱ[color=red]���������ǣ�[/color]"&chr(10)
content=content & phtContent &chr(10)
content=content & "[align=center][URL=z_visual_photo2.asp?act=1&app=1&uid="&getid2&"&pid="&pid&"][color=blue][b][u]�������ͬ���Ӱ[/u][/b][/color][/URL][/align]"&chr(10)
content=content & "[align=center][URL=z_visual_photo2.asp?act=1&app=0&uid="&getid2&"&pid="&pid&"][color=red][b][u]�������ܾ���Ӱ[/u][/b][/color][/URL][/align]"
call messto(myname2,"�������","�����Ӱ",content)
end if	
response.write "��ϲ�����������Ѿ��ɹ�����!" & vbCrLf
case else '�ҵĿ���
        
        dim getmoney
        getmoney=Cint(forum_user(3))
        if mymoney<getmoney then
        response.write"<Br>"+"<li>������֧����д���"&getmoney&"Ԫ���ã���ȥ����^_^��"
 	exit sub
	end if
	if getmyv(Userid)="" then
	response.write"<Br>"+"<li>����û������������ô����˼��������^_^"
   	exit sub
	end if
	
		if DateDiff("s",session("lastpost"),Now())<120 then
   		response.write"<Br>"+"<li>��������������д�棬����ʱ��Ϊ120�룬���Ժ����ġ�"
   		exit sub
		end if
	
        dim mym
	mym=mymoney-getmoney
	conn.execute("update [User] set userWealth="&mym&" WHERE Userid="&userid&"")
set prs=server.createobject("ADODB.Recordset")
sql="select * from [Visual_photo2]"	
prs.open sql,conn,1,3
        prs.addnew
	prs("myid")=userid&"||"
	prs("myv")=getmyv(userid)&"���"
	prs("myc")=getmyc(userid)&"||"
	prs("bake")=""
	prs("pass")=""
	prs("vopen")=0
	prs("lb")=1
	prs("pok")=1
	prs("app")=1
        prs("book")="����"&now()&"��д������Ŷ"
	prs.update
prs.close
set prs=nothing
response.write "��ϲ������д��˳�����!" & vbCrLf
session("lastpost")=Now()
set prs=conn.execute("select top 1 id from [Visual_photo2] order by id desc")
pid=prs(0)
call photosee(pid,"")
end select
end sub

sub photosubok(uid,act,app,pid)'����
dim myid1,myapp,myid2,mylb,mym,getmoney
        if uid="" or getmyname(uid)="" then
	response.write"<Br>"+"<li>���ִ��������µ�½��"
	exit sub
        end if
        
        set prs=server.createobject("adodb.recordset")
        sql="select myid,app,lb from [Visual_photo2] where pok=0 and id="&pid
        set prs=conn.execute(sql)
        if prs.eof or prs.bof then
        response.write"<Br>"+"<li>�������Ӱ�Ѿ����˾ܾ������Ѿ���ɡ�"
        exit sub
	else
	myid1=prs("myid")
	myapp=prs("app")
	myid2=split(myid1,"|")
	mylb=prs("lb")
	if myapp<>mylb then
	response.write"<Br>"+"<li>�˺�Ӱ���뻹����δͬ�⣬��ʱ������ɡ�"
	exit sub
	end if
        end if
        prs.close
        set prs=nothing
        
        if myid2(0)<>cstr(userid)  then
        response.write"<Br>"+"<li>�����Ǵ˴κ�Ӱ�ķ����ˣ����ܵ����ɡ�"
   	exit sub
	end if
        getmoney=Cint(forum_user(3))
        if mymoney<getmoney then
        response.write"<Br>"+"<li>������֧���˴˺�Ӱ��"&getmoney&"Ԫ���ã���ʱ�������^_^��"
 	exit sub
	end if
	mym=mymoney-getmoney
	conn.execute("update [Visual_photo2] set pok=1 WHERE id="&pid&"")
	conn.execute("update [User] set userWealth="&mym&" WHERE Userid="&uid&"")
	response.write "��ϲ�����κ�Ӱ�Ѿ����!<br><br>" & vbCrLf
	call photosee(pid,"")

end sub

sub photolist(lst)'�ҵ�����б�
dim myname,myid,myid1,mylb
dim totalrec,mypid,mypas
if lst="" then lst="0"
if uid="" then 
	uid=userid
        myname="��"
        else
myname=getmyname(uid)
end if

%>
<a name=visualtop>
<table border=0 cellspacing=1 cellpadding=0 bgcolor=<%=Forum_Body(27)%> width=98%>
<tr height=25 valign=middle>
<th valign=middle>
<a href="z_Visual_photo2.asp?act=9&uid=<%=uid%>&lst=3"><%=myname%>�ĺ�Ӱ��</a> -  
<a href="z_Visual_photo2.asp?act=9&uid=<%=uid%>&lst=0"><%=myname%>�ļ�ͥ��</a> - 
<a href="z_Visual_photo2.asp?act=9&uid=<%=uid%>&lst=1"><%=myname%>��д����</a> -  
</th></tr>
<tr height=25 valign=middle>
<td valign=middle width=100%> 
<%
                        set prs=server.createobject("ADODB.Recordset")
                        select case lst
                        case "0" 
			%><%=myname%>�ļ�ͥ��<%
			sql="select id,myid,book,date,pass,cou,lb from [Visual_photo2] where lb=0 and pok=1  order by cou desc"			        
			case "1" 
			%><%=myname%>��д����<%
			sql="select id,myid,book,date,pass,cou,lb from [Visual_photo2] where lb=1 and pok=1  order by cou desc"
			case else '
			%><%=myname%>�ĺ�Ӱ��<%
			sql="select id,myid,book,date,pass,cou,lb from [Visual_photo2] where (lb=2 or lb=3) and pok=1  order by cou desc"
			end select
                        prs.open sql,conn,1,1
                        if prs.eof and prs.bof then
		               response.write "<br>��ʱ��û���κ��û�����</td>"
                        else
                        %></td></tr><tr><td width="100%">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr valign=middle>
			<% 
			                totalrec=prs.recordcount
                                        
					curcount=0
					dim iic
					iic=0
					if curpage > pagecount then curpage = pagecount
						if curpage < 1 then curpage=1
						prs.PageSize=10
						prs.AbsolutePage=CurPage
					do while not prs.eof and curcount<10
				        mylb=prs("lb")
					myid=prs("myid")
					myid1=split(myid,"|")
					mypas=prs("pass")
					dim showok
					
					if myid1(0)=cstr(uid) or myid1(1)=cstr(uid) or myid1(2)=cstr(uid) then
                                        	showok=true
                                        	else
                                        	showok=false
                                        end if
                                        if showok=true then
			                %>
					<td class=tablebody1 width=45%>
					<table width=100% border=0 cellspacing=1 cellpadding=0><tr><td align=center>
					<%
					
			                mypid=prs("id")
					call photosee(mypid,"")
					%>					
                                        </td></tr>
                                        <tr><td class=tablebody2>
                                        <%
                                        response.write "&nbsp;&nbsp;<a href=""z_Visual_photo2.asp?act=4&lst=1&pid="&mypid&"&uid=""&userid&"" title=""����������󣬸���һ���ʻ�����������"&GroupSetting(47)&"���Ǯ""><img src="&Forum_info(7)&Forum_BoardPic(6)&" border=0>�ʻ�</a>&nbsp;&nbsp;<a href=""z_Visual_photo2.asp?act=4&lst=0&pid="&mypid&"&uid=""&userid&"" title=""����������󣬸���һ����������������"&GroupSetting(47)&"���Ǯ""><img src="&Forum_info(7)&Forum_BoardPic(7)&" border=0>����</a>(����<font color="&Forum_body(8)&">"&prs("cou")&"</font>)"
                                        response.write "&nbsp;&nbsp;<a href=""z_Visual_photo2.asp?act=5&lst=0&pid="&mypid&"&uid=""&userid&"" title=""ɾ���⸱��Ӱ"">ɾ��</a>"
                                        %>
                                        </td></tr>
                                        <tr><td class=tablebody2>&nbsp;&nbsp;��Ӱ���ڣ�
                                        <%=prs("date")%><br>&nbsp;&nbsp;
                                        <%=prs("book")%>
                                        </td></tr>
                                        <tr><td class=tablebody2>
                                        <form action="z_Visual_photo2.asp?act=5&pid=<%=mypid%>&lst=1&uid=<%=userid%>&pas=<%=mypas%>" name="subsee" method=post>
                                        <INPUT name=phtname type=text> &nbsp;<input type=submit name=submit value="�Ƽ�������">
                                        </form>
                                        </td></tr>
                                        </table></td>
                                        <%
                                        curcount=curcount+1
                                        iic=iic+1
                                        else
					totalrec=totalrec-1
				        end if
					
					
                                        
                                        if iic mod 2=0 then
                                        iic=0%>
                                        <tr valign=middle>
                                        <%end if
					prs.movenext
					loop
					if totalrec mod 10=0 then
						pagecount=totalrec \ 10
					else
						pagecount=totalrec \ 10+1
					end if
			%>
			</tr></table></td>
			<%		
			end if
			%>
			</tr><tr><td width=100%>
			<table width=100% border=0 cellspacing=0 cellpadding=0>
					<tr height=25 valign=middle>
						<td valign=middle class=tablebody1>&nbsp;&nbsp;ҳ�Σ�<b><%=curPage%></b> / <b><%=pagecount%></b> ҳ ÿҳ��<b>10</b> ���ƣ�<b><%=totalrec%></b></td>
						<td valign=middle class=tablebody1><div align=right>��ҳ��<%
							call DispPageNum(curpage,pagecount,"?act=3&lst="&lst&"&page=","#VisualTop")
						%>&nbsp;&nbsp;</td>
					</tr>
			</table>
			</td></tr></table>
			<%
prs.close
set prs=nothing
end sub


sub photolook(lb,pwidth,mybake,myname1,myname2,myname3,mysex1,mysex2,mysex3,myv1,myv2,myv3,myc1,myc2,myc3,mydate)'��ʾ��Ƭ
dim myfqsex1,myfqsex2,mycolor1,mycolor2,mycolor3
dim amyname,amydate
if lb<>0 then
myfqsex1=""
myfqsex2=""
else '����
if mysex1<>0 then
myfqsex1="[��]"
myfqsex2="[��]"
else
myfqsex2="[��]"
myfqsex1="[��]"
end if	
END IF	



amyname="<table align=left valign=top></tr><td>"
if myname1<>"" then
	if mysex1=1 then
	mycolor1="ffff00" 
	else
	mycolor1="00ff00"
end if
myname1="<table style='filter:glow(color=#05050A,strength=2)'><tr><td align=left><FONT color=#"&mycolor1&">"&myname1&""&myfqsex1&"</font><br><FONT color=#e1e1e1>"&myc1&"</font></td></tr>"
amyname=amyname&myname1&"</table>"
end if

if myname2<>"" then
	if mysex2=1 then
	mycolor2="ffff00" 
	else
	mycolor2="00ff00"
end if
myname2="<table style='filter:glow(color=#01010c,strength=2)'><tr><td align=left><FONT color=#"&mycolor2&">[R]"&myname2&""&myfqsex2&"</font><br><FONT color=#e1e1e1>"&myc2&"</font></td></tr>"
amyname=amyname&myname2&"</table>"
else
myname2=""
end if

if myname3<>"" then
	if mysex3=1 then
	mycolor3="ffff00" 
	else
	mycolor3="00ff00"
end if
myname3="<table style='filter:glow(color=#01010c,strength=2)'><tr><td align=left><FONT color=#"&mycolor3&">[L]"&myname3&"</font><br><FONT color=#e1e1e1>"&myc3&"</font></td></tr>"
amyname=amyname&myname3&"</table>"
else
myname3=""
end if
amyname=amyname&"</td></tr></table>"

mydate="<table style='filter:glow(color=#000000,strength=2)' align=center><tr><td align=left><FONT color=#66FF33>"&mydate&"</font></td></tr></table>"
amydate="<table align=left valign=top></tr><td>"&mydate&"</td></tr></table>"
select case lb
case 3 '���˺�Ӱ
%>
<DIV style='padding:0;position:relative;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/280)%>'>
<img src=<%=mybake%> border=0 style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/280)%>;z-index:-5;'>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-1;'>
<%response.write amyname%>
</DIV>
<DIV style='padding:0;position:absolute;top:<%=int(pwidth*200/280)%>;right:-5;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-1;'>
<%response.write amydate%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:<%=int(pwidth*(1-70/140)/2)%>;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-4;'>
<%
call showphoto(myv3,int(pwidth*140/280),mysex3,-int(pwidth*(1-208/280)/2),"",0)
%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:<%=int(pwidth*(1-70/140)/2)%>;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-2;'>
<%
call showphoto(myv1,int(pwidth*140/280),mysex1,0,"",0)
%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:<%=int(pwidth*(1-70/140)/2)%>;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-3;'>
<%
call showphoto(myv2,int(pwidth*140/280),mysex2,int(pwidth*(1-208/280)/2),"",0)
%>
</DIV>
</DIV>
<%
case 2 '����
if mybake="" then
pwidth=140     'û��ѡ�񱳾�
%>
<DIV style='padding:0;position:relative;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>'>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-1;'>
<%response.write amyname%>
</DIV>
<DIV style='padding:0;position:absolute;top:<%=int(pwidth*200/140)%>;left:10;width:<%=int(pwidth*226/140)%>;height:<%=int(pwidth*226/280)%>;z-index:-1;'>
<%response.write amydate%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-3;'>
<%
call showphoto(myv2,int(pwidth*230/280),mysex2,int(pwidth*(10/280)/2),"",5)
%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-4;'>
<%
call showphoto(myv1,pwidth,mysex1,-int(pwidth*(70/280)/2),"",1)
%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-2;'>
<%
call showphoto(myv1,int(pwidth*230/280),mysex1,-int(pwidth*(150/280)/2),"",5)
%>
</DIV>
</DIV>
<%                              
else
%>
<DIV style='padding:0;position:relative;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/280)%>'>
<img src=<%=mybake%> border=0 style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/280)%>;z-index:-4;'>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-1;'>
<%response.write amyname%>
</DIV>
<DIV style='padding:0;position:absolute;top:<%=int(pwidth*200/280)%>;right:-10;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-1;'>
<%response.write amydate%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:<%=int(pwidth*(1-70/140)/2)%>;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-3;'>
<%
call showphoto(myv1,int(pwidth*140/280),mysex1,-int(pwidth*(1-220/280)/2),"",0)
%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:<%=int(pwidth*(1-70/140)/2)%>;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-2;'>
<%
call showphoto(myv2,int(pwidth*140/280),mysex2,int(pwidth*(1-220/280)/2),"",0)
%>
</DIV>
</DIV>
<%                                	
end if
case 1 '����
pwidth=140
%>
<DIV style='padding:0;position:relative;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>'>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-1;'>
<%response.write amyname%>
</DIV>
<DIV style='padding:0;position:absolute;top:<%=int(pwidth*200/140)%>;left:10;width:<%=int(pwidth*226/140)%>;height:<%=int(pwidth*226/280)%>;z-index:-1;'>
<%response.write amydate%>
</DIV>

<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-3;'>
<%
call showphoto(myv1,pwidth,mysex1,-int(pwidth*(50/280)/2),"",1)
%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-2;'>
<%
call showphoto(myv1,pwidth,mysex1,-int(pwidth*(50/280)/2),"",0)
%>
</DIV>
</DIV>
<%
case 0 '����
if mybake="" then
pwidth=140     'û��ѡ�񱳾�
%>
<DIV style='padding:0;position:relative;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>'>
<DIV style='padding:0;position:relative;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>'>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-1;'>
<%response.write amyname%>
</DIV>
<DIV style='padding:0;position:absolute;top:<%=int(pwidth*200/140)%>;left:10;width:<%=int(pwidth*226/140)%>;height:<%=int(pwidth*226/280)%>;z-index:-1;'>
<%response.write amydate%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-3;'>
<%
call showphoto(myv2,int(pwidth*230/280),mysex2,int(pwidth*(10/280)/2),"",5)
%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-4;'>
<%
call showphoto(myv1,pwidth,mysex1,-int(pwidth*(70/280)/2),"",1)
%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/140)%>;z-index:-2;'>
<%
call showphoto(myv1,int(pwidth*230/280),mysex1,-int(pwidth*(150/280)/2),"",5)
%>
</DIV>
</DIV>
<%                              
else
%>
<DIV style='padding:0;position:relative;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/280)%>'>
<img src=<%=mybake%> border=0 style='padding:0;position:absolute;top:0;left:0;width:<%=pwidth%>;height:<%=int(pwidth*226/280)%>;z-index:-4;'>
<DIV style='padding:0;position:absolute;top:0;left:0;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-1;'>
<%response.write amyname%>
</DIV>
<DIV style='padding:0;position:absolute;top:<%=int(pwidth*200/280)%>;right:-10;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-1;'>
<%response.write amydate%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:<%=int(pwidth*(1-70/140)/2)%>;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-2;'>
<%
call showphoto(myv1,int(pwidth*140/280),mysex1,-int(pwidth*(1-240/280)/2),"",0)
%>
</DIV>
<DIV style='padding:0;position:absolute;top:0;left:<%=int(pwidth*(1-70/140)/2)%>;width:<%=int(pwidth*120/280)%>;height:<%=int(pwidth*226/280)%>;z-index:-3;'>
<%
call showphoto(myv2,int(pwidth*140/280),mysex2,int(pwidth*(1-240/280)/2),"",0)
%>
</DIV>
</DIV>
<%
end if
end select
end sub

sub showphoto(uservisual,visualwidth,userstex,visualtop,username,inchtc)
 response.write "<table border=0 cellspacing=0 cellpadding=0 width="&visualwidth&">"
 response.write "<tr>"
 response.write "<td align=center visualwidth="&visualwidth&"  visualheight="&(visualwidth*226/140)&" ImagePath="""&PicPath&""" username="""&username&""" visualtop="&visualtop&" usergender="&userstex&" baseName=""images/img_visual/blank.gif"" visual="""&uservisual&"""  style=behavior:url(inc/z_show"&inchtc&".htc) localpic="""&LocalPic&""" ></td>"
 response.write "</tr>"
 response.write "</table>"
end sub


sub messto(towho,sender,title,content)
    conn.execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&towho&"','"&sender&"','"&title&"','"&content&"',Now(),0,1)")
end sub

'����û�����
function getmyname(ismyid)
dim mynametmp
dim mrs
mynametmp=""
if ismyid="" then
	getmyname=""
	Exit Function
	else
	sql="select UserName from [user] where UserID="&ismyid
	set mrs=conn.execute(sql)
        if mrs.eof or mrs.bof then
        	getmyname=""
	Exit Function
	else
	mynametmp=mrs("UserName")
end if
mrs.close
set mrs=nothing	
end if
getmyname=mynametmp
end function

'����û��Ա�
function getmysex(ismyid)
dim mysextmp
dim mrs
mysextmp=""
if ismyid="" then
	getmysex=""
	Exit Function
	else
	sql="select Sex from [user] where UserID="&ismyid
	set mrs=conn.execute(sql)
        if mrs.eof or mrs.bof then
        	getmysex=""
	Exit Function
	else
	mysextmp=mrs("Sex")
end if
mrs.close
set mrs=nothing	
end if
getmysex=mysextmp
end function

'����û�ID
function getmyid(ismyname)
dim myidtmp
dim mrs
myidtmp=""
if ismyname="" then
	getmyid=""
	Exit Function
	else
	sql="select UserID from [user] where UserName='"&ismyname&"'"
	set mrs=conn.execute(sql)
        if mrs.eof or mrs.bof then
        	getmyid=""
	Exit Function
	else
	myidtmp=mrs("UserID")
end if
mrs.close
set mrs=nothing	
end if
getmyid=myidtmp
end function

'ͨ��ID����û���������
function getmyv(ismyid)
dim myvtmp
dim mrs
myvtmp=""
        if ismyid="" then
	getmyv=""
	Exit Function
	else
	sql="select visual from [user] where UserID="&ismyid
	set mrs=conn.execute(sql)
        if mrs.eof or mrs.bof then
        	getmyv=""
	Exit Function
	else
	myvtmp=mrs("visual")
	if myvtmp="" or isnull(myvtmp) then
	getmyv=""
	end if
end if
mrs.close
set mrs=nothing	
end if
getmyv=myvtmp
end function

'ͨ��ID����û��ȼ�
function getmyc(ismyid)
dim myctmp
dim mrs
myctmp=""
        if ismyid="" then
	getmyc=""
	Exit Function
	else
	sql="select userclass from [user] where UserID="&ismyid
	set mrs=conn.execute(sql)
        if mrs.eof or mrs.bof then
        	getmyc=""
	Exit Function
	else
	myctmp=mrs("userclass")
end if
mrs.close
set mrs=nothing	
end if
getmyc=myctmp
end function

function getmyw(vismywife,vismyid)
dim mrs,ismywife
        if vismywife="" or vismyid="" then
	getmyw=false
	else
	sql="select UserID from [user] where UserID="&vismyid&" and wife='"&vismywife&"'"
	set mrs=conn.execute(sql)
        if mrs.eof or mrs.bof then
        	getmyw=false
	else
	getmyw=true
	end if
        end if
mrs.close
set mrs=nothing	
end function
%>