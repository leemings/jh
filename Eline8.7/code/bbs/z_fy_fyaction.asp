<!--#include file="conn.asp"-->
<!--#include file="z_fy_conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
 response.buffer=true
'=========================================================
' Plusname: ��Ժ����������
' File: z_fy_fyaction.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
stats="��Ժ����ͥ"
call nav()
call head_var(0,0,"������Ժ","z_fy_fayuan.asp")
call getfyconfig()
if not master and not fymaster and not jymaster and request("faction")<>"baoshi" and request("faction")<>"ceshu" then
  Errmsg=Errmsg+"<br>"+"<li>��û�н��뱾������Ժ����ͥ��Ȩ�ޣ���ʹ���˷Ƿ�������"
  call dvbbs_error()
else
  response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>��Ժ����ͥ</b></td></tr><tr><td valign=middle class=tablebody2 height=100><CENTER>"
  dim id,fyrs,fmx,fbgxx,fygxx,fmz,result,bg,my,fgc,fgp,fysql,fjtype
  id=request("id")
  select case checkstr(request("faction"))
    case "alldel"
      call alldel()
    case "allout"
      call allout()
    case "baoshi"
      call baoshi()
       fyrs.close
       set fyrs=nothing
    case "ceshu"
      if not isInteger(request("id")) then response.end
      call ceshu()
    case "peishen"
      if not isInteger(request("id")) then response.end
       call peishen()
    case else
      result=checkstr(request("result"))
      bg=checkstr(request("bg"))
      my=checkstr(request("my"))
      if isnull(request.form("panci")) or checkstr(request.form("panci"))="��" then
        fgc="δ������"
      else
        fgc=checkstr(request.form("panci"))
      end if
      fgp=checkstr(request.form("play1"))
      id=checkstr(request.querystring("id"))
      if not isInteger(id) then response.end
      select case checkstr(Request.form("action"))
        case ""
           main()
           fyrs.close
           set fyrs=nothing
        case "agree"
           fmz="1"
           agree()
        case "noagree"
           noagree()
        case "pingfan"
           pingfan()
        case "wugao"
           wugao()
        case "wugaohigh"
           wugaohigh()
        case "del"
           del()
        case "fgagree"
           result=fgp
           fmz="6"
           agree()
      end select
  end select
  connfy.close
  set connfy=nothing
  response.write "</CENTER></td></tr></table>" 
end if
call footer() 
'=========================================================
' Sub: main
' Readme:���ٵ���������ְ��
'=========================================================
Sub main()
fysql="SELECT * FROM fy WHERE ID=" & id
set fyrs=connfy.execute(fysql)
if fyrs.eof or fyrs.bof then
  Errmsg=Errmsg+"<br>"+"<li>û���ҵ������״��"
  call dvbbs_error()
else
  if fymaster and fyrs(7)<>"N" then
    Errmsg=Errmsg+"<br>"+"<li>�Բ�����������û������Ȩ����"
    call dvbbs_error()
    exit sub
  end if
  response.write "<div align=center><br><b>�������<font color=red>"&id&"</font> </b><br></div>"
  response.write "<table width=97% border=0 cellspacing=1 cellpadding=3 align=center class=tableborder1>"
  response.write "<tr><td height=24 class=tablebody2>ԭ  �� <a href=dispuser.asp?name="&fyrs(4)&" target=_blank><b>"&fyrs(4)&"</b></a>&nbsp;��ŭ��д����"&fyrs(1)&"</td></tr>"
  response.write "<tr><td class=tablebody1><br>"&checkstr(fyrs(3))&"<br><br>���� <a href=dispuser.asp?name="&left(fyrs(2),10)&" target=_blank><b>"&left(fyrs(2),10)&"</b></a> ���ֶ����о���ϣ������ <b><font color=#FF0000>"&fyrs(5)&"</font></b> �Ĵ�����</td></tr>"
  response.write "<tr><td height=24 class=tablebody2>��  �� <a href=dispuser.asp?name="&fyrs(2)&" target=_blank><b>"&fyrs(2)&"</b></a>&nbsp;������ </td></tr>"
  response.write "<tr><td class=tablebody1><br>"&checkstr(fyrs(9))&"<br></td></tr>"
  response.write "<tr><td height=24 align=left class=tablebody2><form action=z_fy_fyaction.asp?result="&fyrs(5)&"&id="&id&"&bg="&fyrs(2)&"&my="&fyrs(4)&" method = post> ���� <b>"&membername&"</b> �о����£�</td></tr>"
  response.write "<tr><td height=24 align=left class=tablebody1>���ԣ�<input name=panci type=text size=80 value="&fyrs(10)&">"
  select case  fyrs(7)
       case "N" 
         response.write "�����з�"
       case "B"
         response.write "����ԭ�����"
       case "P"
         response.write "ԭ����ƽ��"
       case "C"
         response.write "ԭ���ѳ���"
       case else
         response.write "�ѽ��� <font color=red><b>"&fyrs(7)&"</b></font> �Ĵ���"
  end select
  response.write "</td></tr>"
  response.write "<tr><td height=24 align=left class=tablebody2><b>�о�ѡ�</b>"
  response.write "<select name=action size=1><option value='agree'><b>ͬ��ԭ��</b></option><option value='noagree'><b>��������</b></option>"
  response.write "<option value='pingfan'><b>����ƽ��</b></option><option value='wugao'><b>ԭ���ܸ�</b></option>"
  response.write "<option value='wugaohigh'><b>�����ܸ�</b></option><option value='del'><b>ɾ����״</b></option><option value='fgagree'><b>���������̡�</b></option></select>&nbsp;&nbsp;"
  response.write "<b>�������̳߶ȣ�</b><select name='play1' size=1>"
  %>
                <option value="�������"><b>���������</b></option>
                <option value="����10����"><b>&nbsp;&nbsp;������10����</b></option>
                <option value="������Сʱ"><b>&nbsp;&nbsp;��������Сʱ</b></option>
                <option value="����һСʱ"><b>&nbsp;&nbsp;������һСʱ</b></option>
                <option value="����һ��"><b>&nbsp;&nbsp;������һ��</b></option>
                <option value="��������"><b>&nbsp;&nbsp;����������</b></option>
                <option value="����һ��"><b>&nbsp;&nbsp;������һ��</b></option>
                <option value="û��ȫ���Ʋ�"><b>��Է���ȫ���Ʋ���</b></option>               
                <option value="����1��"><b>&nbsp;&nbsp;������1%</b></option>
                <option value="����10��"><b>&nbsp;&nbsp;������10%</b></option>
                <option value="����50��"><b>&nbsp;&nbsp;������50%</b></option>
                <option value="�۳����о���"><b>��Կ۳����о����</b></option>               
                <option value="�۾���1��"><b>&nbsp;&nbsp;���۾���1%</b></option>
                <option value="�۾���10��"><b>&nbsp;&nbsp;���۾���10%</b></option>               
                <option value="�۾���50��"><b>&nbsp;&nbsp;���۾���50%</b></option>
                <option value="�۳���������"><b>��Կ۳�����������</b></option> 
                <option value="������1��"><b>&nbsp;&nbsp;��������1%</b></option>               
                <option value="������10��"><b>&nbsp;&nbsp;��������10%</b></option>               
                <option value="������50��"><b>&nbsp;&nbsp;��������50%</b></option>
                <option value="���"><b>������з�ֵ�����</b></option> 
                <option value="����-10"><b>&nbsp;&nbsp;������-10</b></option>
                <option value="����-50"><b>&nbsp;&nbsp;������-50</b></option>
                <option value="����-100"><b>&nbsp;&nbsp;������-100</b></option>
                <option value="����ն��"><b>�������ն�ס�</b></option>
  <%
  if lh_set(0)="1" then response.write "<option value='���'><b>����뱻������</b></option>"
  response.write "</select>&nbsp;&nbsp;<input type=checkbox name='giveto'  value='1'>��û���ָ�ʤ��(�����з��ۼ�����ʱ��Ч��"
  response.write "</td></tr><tr><td class=tablebody1 align=left><b>����ѡ�</b><input type=checkbox name='nobaoshi'  value='1'>��ֹ����&nbsp;&nbsp;�Զ����ͽ�:<input name=bsmoney type=text size=10 value="&bs_set(1)&"> �����Ͻ����з�����ʱ��Ч)</td></tr>"
  response.write "<tr><td class=tablebody2 height=24 align=center><input type=submit name=Submit value='ִ ��'>&nbsp;&nbsp;<input type=button value='�� ��' onClick='javascript:history.back()' name=button></td></tr></table>"
end if
End sub
'=========================================================
' Sub: del
' Readme:������������-ɾ������״ֽ
'=========================================================
Sub del()
  sql="delete from fy where id=" & id & ""
  set fyrs=connfy.execute(sql)
  'fyrs.close
  set fyrs=nothing
  Response.Redirect "z_fy_over.asp"
End sub
'=========================================================
' Sub: alldel
' Readme:״ֽ�б���-ɾ������᰸��
'=========================================================
Sub alldel()
if master then
  sql="delete from fy where result<>'N'"
  connfy.execute sql
end if
  Response.Redirect "z_fy_over.asp"
End sub
'=========================================================
' Sub: allout
' Readme:������������-�ͷ����з���
'=========================================================
Sub allout()
if master then
  sql="update [user] set lockuser=0 where lockuser=1"
  conn.execute sql
  fysql="update fy set stats='7',dateandtime=now() where result='��������' or result='����һ��' or result='����һ��' or result='�������'"
  connfy.execute fysql
  call bnew("��������ʵ�д���","�ͷż����й�Ѻ�����з���",membername)
end if
  Response.Redirect "z_fy_fanren.asp"
End sub
'=========================================================
' Sub: baoshi
' Readme:����-����ĳһ����
'=========================================================
Sub baoshi()
dim mysign_set
  if bs_set(0)="0" then
    Errmsg=Errmsg+"<br>"+"<li>��������Ŀǰ�������ͣ�"
    call dvbbs_error()
    exit sub
  end if
  sql="select top 1 sign from log where class='����' and title='����' and tousername='"&request("name")&"' ORDER BY dateandtime DESC"
  set fyrs=connfy.execute(sql)
  if fyrs.eof then
    Errmsg=Errmsg+"<br>"+"<li>û�и÷��˵ļ�¼�����ܱ��ͣ�"
    call dvbbs_error()
    exit sub
  end if
  mysign_set=split(fyrs(0),"|")
  if not jymaster and not master then
    if mysign_set(0)="1" then
       Errmsg=Errmsg+"<br>"+"<li>�÷��˲��ܱ��ͣ�ֻ���ɼ������ͷţ�"
       call dvbbs_error()
       exit sub
    end if
    sql="select userwealth from [user] where username='"&membername&"'"
    set rs=conn.execute(sql)
    if clng(rs(0))<=clng(mysign_set(1)) then
      Errmsg=Errmsg+"<br>"+"<li>���ı��ͽ𲻹�����������"+mysign_set(1)+"Ԫ���ܱ��͸÷���"
      call dvbbs_error()
      exit sub
    end if
    sql="update [user] set userwealth=userwealth-"&bs_set(1)&" where  username='"&membername&"'"
    conn.execute sql
    content="������Ѻ���� "&request("name")&"��������"&mysign_set(1)&"Ԫ"
    mysign="0|0"
    call logs("����","����",membername,request("name"))
  else
    mysign="0|0"
    content="����"&request("name")&"��Ѻ�ڼ�������ã����������ͷ�"
    call logs("����","����",membername,request("name"))
  end if
  sql="update [user] set lockuser=0 where lockuser=1 and username='"&request("name")&"'"
  conn.execute sql
  Response.Redirect "z_fy_fanren.asp"
End sub
'=========================================================
' Sub: ceshu
' Readme:ԭ�泷�ߴ���
'=========================================================
Sub ceshu()
  sql="select yg,result from fy where id=" & id & ""
  set fyrs=connfy.execute(sql)
  if membername<>fyrs(0) or fyrs(1)<>"N" then
    Errmsg=Errmsg+"<br>"+"<li>��û�г�������״��Ȩ�޻�ð��ѽ᰸��"
    call dvbbs_error()
    exit sub
  end if
  fysql="update fy set stats='8',result='C',dateandtime=now() where id=" & id & ""
  connfy.execute(fysql)
  fyrs.close
  set fyrs=nothing
  Response.Redirect "z_fy_over.asp"
End sub
'=========================================================
' Sub: peishen()
' Readme:�ύ����������
'=========================================================
Sub peishen()
	if pst_set(0)="0" then
		Errmsg=Errmsg+"<br>"+"<li>���������鹦��Ŀǰ���رգ�����ʹ����򿪲�������ӦͶƱ���棡"
		call dvbbs_error()
		exit sub
	else
		boardid=trim(pst_set(1))
		fysql="select * from fy where id="&id
		set fyrs=connfy.execute(fysql)
		response.write "<form action=Savevote.asp?boardID="&boardid&" method=POST onSubmit=submitonce(this) name=frmAnnounce>"
		response.write "<input type=hidden name=upfilerename>"
		response.write "<script src='inc/ubbcode.js'></script>"
		response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1>"
		response.write "<tr><th width=100% colspan=2 align=center>&nbsp;&nbsp;�����ֵ� "&id&" �Ű����ͽ�����������"
		if isaudit=1 then response.write "���������������Ӷ�Ҫ��������Ա��˷��ɷ���"
		response.write "</th></tr>"
		response.write "<tr><td width=20% class=tablebody2><b>�û�����</b></td>"
		response.write "<td width=80% class=tablebody2><input name=username value="&membername&" class=FormClass disabled>&nbsp;&nbsp;<b>���룺</b><input name=passwd type=password value="&htmlencode(memberword)&" class=FormClass disabled></td></tr>"
		response.write "<tr><td width=20% class=tablebody1><b>���⣺</b></td>"
		response.write "<td width=80% class=tablebody2><input name=subject size=50 maxlength=80 class=FormClass value='[���ֵ�"&id&"�Ű���]���λ��������'>&nbsp;<select name=votetimeout size=1>"
		response.write "<option value=0>����ʱ��</option><option value=0>��������</option><option value=1>һ��</option><option value=3>����</option><option value=7>һ��</option><option value=15>����</option></select>&nbsp;<b>*</b>���ⲻ�ó���25�����ֻ�50��Ӣ���ַ�"
		response.write "</td></tr>"
		response.write "<tr><td width=20% class=tablebody2><b>�з�ѡ� </b> <br><br><li>ÿ��һ���з���Ŀ</li><li>���ٿɸ��ݾ�������޸�</li><br><br>"
		response.write "<input type=radio name=votetype value=0 checked >��ѡͶƱ<br></font></td>"
		response.write "<td width=80% class=tablebody2><textarea name=vote cols=95 rows=8>ͬ��ԭ�棨��ԭ������������棩"&chr(13)&"ͬ��ԭ�棨���Ա�����ش�����"&chr(13)&"ͬ��ԭ�棨�ɶԱ�����ᴦ����"&chr(13)&"ԭ�洿����ҥ��Ӧ�ܴ���"&chr(13)&"ԭ�洿���ܸ棬Ӧ���ط�</textarea>"
		response.write "</td></tr>"
		response.write "<tr><td width=20% valign=top class=tablebody1><b>��ǰ���飺</b><br><li>���������ӵ�ǰ��</td><td width=80% class=tablebody1>"
		for i=0 to Forum_PostFaceNum-1
			response.write "<input type=radio value='"&Forum_PostFace(i)&"' name='Expression'"
			if i=0 then response.write "checked"
			response.write "><img src="&forum_info(8)+Forum_PostFace(i)&" >&nbsp;&nbsp;"
			if i>0 and ((i+1) mod 9=0) then response.write "<br>"
		next
		response.write "</td></tr>"
		response.write "<tr><td width=20% valign=top class=tablebody2><b>�������飺</b><br></td><td width=80% class=tablebody2><br>"
		response.write "<textarea class=smallarea cols=95 name=Content rows=12 wrap=VIRTUAL title=����ʹ��Ctrl+Enterֱ���ύ���� class=FormClass onkeydown=ctlent()>[b]Ͷ�����ڣ�[/b]"&fyrs(8)&chr(13)
		response.write "[b]���棺[/b]"&fyrs(2)&chr(13)&"[b]ԭ�棺[/b]"&fyrs(4)&chr(13)&"[b]ԭ��Ҫ��Ա��洦��"&fyrs(5)&"�ĳͷ�[/b]"&chr(13)&chr(13)&"[b]Ͷ�߱��⣺[/b]"&fyrs(1)&chr(13)&"[b]Ͷ�����飺[/b]"&fyrs(3)&chr(13)&"[b]�������ߣ�[/b]"&fyrs(9)&chr(13)&chr(13)&"[color=#FF0000][url=z_fy_fyaction.asp?id="&id&"]����������������������[/url][/color]</textarea>"
		response.write "</td></tr>"
		response.write "<tr><td valign=middle colspan=2 align=center class=tablebody2><input type=Submit value='�� ��' name=Submit> &nbsp; <input type=button value='Ԥ ��' name=Button onclick=gopreview()>&nbsp;<input type=reset name=Submit2 value='�� ��'></td></form></tr></table></form>"
		response.write "<form name=preview action=preview.asp?boardid="&boardid&" method=post target=preview_page><input type=hidden name=title value=><input type=hidden name=body value=></form>"
	end if
	fyrs.close
	set fyrs=nothing
End sub
'=========================================================
' Sub: fmcl()
' Readme:��ͨ����-ͬ��ԭ����ִ���
'=========================================================
sub fmcl()
sql="select "&fmx&"*"&fygxx&" from [user] where username='"&bg&"'"
set rs=conn.execute(sql)
if request.form("giveto")="1" then
 sql="update [user] set "&fmx&"="&fmx&"+round("&rs(0)&") where username='"&my&"'"  
 conn.execute sql
end if
fysql="update fy set lastvalue=round("&rs(0)&") where id=" & id & ""
connfy.execute fysql
sql="update [user] set "&fmx&"=round("&fmx&"*"&fbgxx&") where username='"&bg&"'"  
conn.Execute(sql)
end sub
'=========================================================
' Sub: fjfj()
' Readme:���ӷ�����
'=========================================================
sub fjfj()
select case fjtype
case "1"
  if fmz="1" and fjty_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjty_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fmz="1" and fjty_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjty_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
case "2"
  if fjbh_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjbh_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fjbh_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjbh_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
case "3"
  if fjpf_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjpf_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fjpf_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjpf_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
case "4"
  if fjzy_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjzy_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fjzy_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjzy_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
case "5"
  if fjwg_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjwg_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fjwg_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjwg_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
end select
end sub
'=========================================================
' Sub: agree()
' Readme:ͬ��ԭ��/�����Զ����з�
'=========================================================
sub agree()
fjtype="1"
select case result
case "����1��" 
  fmx="userwealth"
  fbgxx="0.99"
  fygxx="0.01"
  call fmcl()
case "����10��" 
  fmx="userwealth"
  fbgxx="0.9"
  fygxx="0.1"
  call fmcl()
case "����50��" 
  fmx="userwealth"
  fbgxx="0.5"
  fygxx="0.5"
  call fmcl()
case "û��ȫ���Ʋ�"
  fmx="userwealth"
  fbgxx="0"
  fygxx="1"
  call fmcl()
case "�۾���1��" 
  fmx="userep"
  fbgxx="0.99"
  fygxx="0.01"
  call fmcl()
case "�۾���10��"
  fmx="userep"
  fbgxx="0.9"
  fygxx="0.1"
  call fmcl()
case "�۾���50��"
  fmx="userep"
  fbgxx="0.5"
  fygxx="0.5"
  call fmcl()
case "�۳����о���" 
  fmx="userep"
  fbgxx="0"
  fygxx="1"
  call fmcl()
case "������1��"
  fmx="usercp"
  fbgxx="0.99"
  fygxx="0.01"
  call fmcl()
case "������10��"
  fmx="usercp"
  fbgxx="0.9"
  fygxx="0.1"
  call fmcl()
case "������50��"
  fmx="usercp"
  fbgxx="0.5"
  fygxx="0.5"
  call fmcl()
case "�۳���������" 
  fmx="usercp"
  fbgxx="0"
  fygxx="1"
  call fmcl()
case "����-10"
  sql="update [user] set userpower=userpower-10 where username='" & bg & "'"
  conn.execute sql
  if request.form("giveto")="1" then
    sql="update [user] set userpower=userpower+10 where username='"&my&"'"  
    conn.execute sql
   end if
case "����-50"
  sql="update [user] set userpower=userpower-50 where username='" & bg & "'"
  conn.execute sql
  if request.form("giveto")="1" then
    sql="update [user] set userpower=userpower+50 where username='"&my&"'"  
    conn.execute sql
   end if
case "����-100" 
  sql="update [user] set userpower=userpower-100 where username='" & bg & "'"
  conn.execute sql
   if request.form("giveto")="1" then
    sql="update [user] set userpower=userpower+100 where username='"&my&"'"  
    conn.execute sql
   end if
case "���" 
  sql="update [user] set userwealth=0,usercp=0,userep=0,userpower=0 where username='" & bg & "'"
  conn.execute sql
case "����10����" 
   content="����Υ�������涨��������10����"
   call jjy()
case "������Сʱ" 
   content="����Υ�������涨����������Сʱ"
   call jjy()
case "����һСʱ" 
   content="����Υ�������涨��������һСʱ"
   call jjy()
case "����һ��" 
   content="����Υ�������涨��������һ��"
   call jjy()
case "��������"  
  content="����Υ�������涨������������"
  call jjy()
case "����һ��" 
  content="����Υ�������涨��������һ��"
  call jjy()
case "�������"
  content="����Υ�������涨�����������"
  call jjy()
case "����ն��" 
  sql="delete from [user] where username='" & bg & "'"
  conn.execute sql
case "���"
  if lh_set(0)="0" then
    Errmsg=Errmsg+"<br>"+"<li>����з�����Ŀǰ�ǹرյģ�"
    call dvbbs_error()
    exit sub
  end if
  if lh_set(0)="1" then
    '���´�����lwand�ṩ
    sql="update [user] set wife=' ' where username='" & bg & "'"' or username='" & my "' "
    conn.execute(sql)
    connj.execute("delete from jie where username='" & my & "'")
    connj.execute("delete from jie where username='" & bg & "'")
    connj.execute("delete from qiuhun where username='" & my & "'")
    connj.execute("delete from qiuhun where username='" & bg & "'")
    '���ϴ�����lwand�ṩ
    if lh_set(1)="1" then
      sql="update [user] set usercp=usercp-"&lh_set(2)&" where username='"&bg&"'"  
      conn.execute(sql)
    end if
  end if
end select
if founderr then 
 call dvbbs_error()
 exit sub
end if
call fjfj()
fysql="update fy set stats='"&fmz&"',fgtext='"&fgc&"',result='"&result&"',dateandtime=now() where id=" & id & ""
connfy.execute fysql
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: noagree()
' Readme:����ԭ��
'=========================================================
sub noagree()
fjtype="2"
fysql="update fy set stats='0',fgtext='"&fgc&"',result='B',dateandtime=now() where id=" & id & ""
response.write fysql
connfy.execute fysql
call fjfj()
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: pingfan()
' Readme:ƽ��
'=========================================================
sub pingfan()
fjtype="3"
call getlast()
if fyrs(0)="����1��" or fyrs(0)="����10��" or fyrs(0)="����50��" or fyrs(0)="û��ȫ���Ʋ�" then
  sql="update [user] set userwealth=userwealth+"&fyrs(1)&" where username='"&bg&"'"  
  conn.Execute(sql)
end if
if fyrs(0)="�۾���1��" or fyrs(0)="�۾���10��" or fyrs(0)="�۾���50��" or fyrs(0)="�۳����о���" then
    sql="update [user] set userep=userep+"&fyrs(1)&" where username='" & bg & "'"
    conn.execute sql
end if
if fyrs(0)="������1��" or fyrs(0)="������10��" or fyrs(0)="������50��" or fyrs(0)="�۳���������" then
    sql="update [user] set usercp=usercp+"&fyrs(1)&" where username='" & bg & "'"
    conn.execute sql
end if
if fyrs(0)="����-10" then
sql="update [user] set userpower=userpower+10 where username='" & bg & "'"
conn.execute sql
end if
if fyrs(0)="����-50" then
sql="update [user] set userpower=userpower+50 where username='" & bg & "'"
conn.execute sql
end if
if fyrs(0)="����-100" then
sql="update [user] set userpower=userpower+100 where username='" & bg & "'"
conn.execute sql
end if
if fyrs(0)="���" or fyrs(0)="N" or fyrs(0)="P" or fyrs(0)="����ն��" or fyrs(0)="C" or fyrs(0)="���" then 
   Errmsg=Errmsg+"<br>"+"<li>�ð�����δ���л��ѳ��߻���о��޷�ƽ����"
   call dvbbs_error()
   exit sub
end if
if fyrs(0)="����10����" or fyrs(0)="������Сʱ" or fyrs(0)="����һСʱ" or fyrs(0)="����һ��" or fyrs(0)="��������" or fyrs(0)="����һ��" or fyrs(0)="�������" then
   sql="update [user] set lockuser=0 where username='" & bg & "'"
   conn.execute sql
   mysign="0|0"
   content="������ƽ�������Գ���"
   call logs("����","����",membername,bg)
end if
if fyrs(0)="B" then
  sql="update [user] set userwealth=userwealth+"&fyrs(1)&" where username='"&my&"'"  
  conn.Execute(sql)
end if
fysql="update fy set stats='2',fgtext='"&fgc&"',result='P',dateandtime=now() where id=" & id & ""
connfy.execute fysql
call fjfj()
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: wugao()
' Readme:�ܸ�
'=========================================================
sub wugao()
fjtype="4"
call fjfj()
fysql="update fy set stats='4',fgtext='"&fgc&"',result='B',dateandtime=now(),lastvalue="&fjzy_set(0)&" where id=" & id & ""
connfy.execute fysql
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: wugaohigh()
' Readme:�����ܸ�
'=========================================================
sub wugaohigh()
fjtype="5"
call fjfj()
fysql="update fy set stats='5',fgtext='"&fgc&"',result='B',dateandtime=now(),lastvalue="&fjzy_set(0)&" where id=" & id & ""
connfy.execute fysql
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: getlast()
' Readme:��ȡ�ϴα������֣�����ƽ��
'=========================================================
sub getlast()
fysql="select result,lastvalue from fy where id=" & id & ""
set fyrs=connfy.execute(fysql)
end sub
'=========================================================
' Sub: jjy()
' Readme:��������
'=========================================================
sub jjy()
dim bbsmoney
if not isInteger(request.form("bsmoney")) or request.form("bsmoeny")<0 then
   Errmsg=Errmsg+"<br>"+"<li>���趨�ı��ͽ𲻺Ϸ���"
   founderr=true
   exit sub
else
if clng(request.form("bsmoney"))<=clng(bs_set(1)) then
  bbsmoney=bs_set(1)
else
  bbsmoney=request.form("bsmoney")
end if
  sql="update [user] set lockuser=1 where username='" & bg & "'"
  conn.execute sql
  if request.form("nobaoshi")="1" then
    mysign="1" & "|" & bbsmoney
  else
    mysign="0" & "|" & bbsmoney
  end if
  call logs("����","����",membername,bg)
end if
end sub
%>
