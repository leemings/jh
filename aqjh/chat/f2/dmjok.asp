<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "../chaterr.asp?id=001" 
end if 

nickname=Session("aqjh_name")
grade=Session("aqjh_grade")
chatroomsn=session("nowinroom")

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
id=LCase(trim(request.querystring("id")))
fromname=trim(request.querystring("from1"))
to1=trim(request.querystring("to1"))
yn=trim(request.querystring("yn"))

if to1="���" or to1=fromname then call ErrALT("�㲻����ѡ���һ�������Ϊ���֣�")

if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End
end if
if nickname<>to1 then
	msg="�˼�û��������ѽ��"
	abc=0
elseif yn=0 then
	msg="��ܾ��˼ҵĶԶ��ˣ�"
	abc=1
	duidu="���ܾ���["&nickname&"]�ܾ�["& fromname &"]�Ĵ��齫����!"
	duidu=duidu & "<br>.........���齫�Ѹ�ý����һ�Ҫ�о�һ��...."
else
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	db="DMJ.ASP"
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.open connstr
	sql="select * from dmj where ufrom='"& nickname & "' or ufrom='"& fromname & "' or uto='"& nickname &"' or uto='"& fromname &"'"
	Set Rs=conn.Execute(sql)
	if not (rs.eof and rs.bof) then
	  if rs("ufrom")=nickname or rs("ufrom")=fromname then
		  unpkname=rs("ufrom")
	  else
	  	  unpkname=rs("uto")
	  end if
	  rs.close
	  set rs=nothing
	  conn.close
	  set conn=nothing
	  response.write "<script Language='Javascript'>alert('" & unpkname & "�����ƾ�û�н����������ٴο���')</script>"
	  abc=1
	  duidu="��������["&nickname&"]���ܽ���["& fromname &"]�ĶԶ�����!,��Ϊ["&unpkname&"]�����������ƾ�û�н�����"
	  msg=unpkname&"�����ƾ�û�н����������ٴο��֣�"
	else
	  sql="select duz from dmj where uto='"& nickname & "$��ע' and ufrom='"& fromname & "$��ע' order by id Desc"
	  Set Rs=conn.Execute(sql)
	  If rs.eof and rs.bof Then
	  		Response.Write "<script language=JavaScript>{alert('����û������ľ֣�');}</script>"
			rs.close
			conn.close
			set rs=nothing
			set conn=nothing
			response.end
	  end If
	  xiazhu=Rs(0)
	  sql="delete * from dmj where uto='"& nickname & "$��ע' or ufrom='"& fromname & "$��ע' or ufrom='"& nickname & "$��ע' or uto='"& fromname & "$��ע'"
	  Set Rs=conn.Execute(sql)
	  dim xia_class,xia_fir,xia_sql,duidupd
		xia_fir=left(xiazhu,1)
		xiazhu1=mid(xiazhu,2)
		select case xia_fir
			case "1"
				xia_class="���"
				xia_sql="���"
			case "2"
				xia_class="����"
				xia_sql="����"
			case "3"
				xia_class="����"
				xia_sql="�ݶ�����"
			case else
				call ErrALT("�Ƿ�������")
		end select
		duidupd=0
		Set conn1=Server.CreateObject("ADODB.CONNECTION")
		Set rs1=Server.CreateObject("ADODB.RecordSet")
		conn1.open application("aqjh_usermdb")
		sql="select " & xia_sql & " from �û� where ����='"& nickname &"'"
		set rs1=conn1.execute(sql)
		sjxiazu=rs1(0)
		rs1.close
		if sjxiazu<(clng(xiazhu1)*2) then
			msg=nickname&"�Ķ�ע"&xia_class&"���㣬���ܴ��齫��Ҫ�����¶�ע�������ϵ��ʽ���ܴ��齫��"
			abc=1
			duidu=nickname&"�Ķ�ע"&xia_class&"���㣬���ܴ��齫��Ҫ�����¶�ע�������ϵ��ʽ���ܴ��齫��"
			duidupd=1
		end if
		myxiazu=sjxiazu-clng(xiazhu1)*2
		sql="select " & xia_sql & " from �û� where ����='"& fromname &"'"
		set rs1=conn1.execute(sql)
		sjxiazu=rs1(0)
		rs1.close
		if sjxiazu<(clng(xiazhu1)*2) then
			msg=fromname&"�Ķ�ע"&xia_class&"���㣬���ܴ��齫��Ҫ�����¶�ע�������ϵ��ʽ���ܴ��齫��"
			abc=1
			duidu=fromname&"�Ķ�ע"&xia_class&"���㣬���ܴ��齫��Ҫ�����¶�ע�������ϵ��ʽ���ܴ��齫��"
			duidupd=1
		end if
		toxiazu=sjxiazu-clng(xiazhu1)*2
		if duidupd=0 then	'�۳����¶Ľ��2����������ʤ�����ٷ���
			sql="update �û� set "& xia_sql &"="& myxiazu &" where ����='"& nickname &"'"
			set rs1=conn1.execute(sql)
			sql="update �û� set "& xia_sql &"="& toxiazu &" where ����='"& fromname &"'"
			set rs1=conn1.execute(sql)
		else
			sql="delete * from mjInfo where ׯ��='" & nickname & "' or ׯ��='" & fromname & "' or ����='" & nickname & "' or ����='" & fromname & "'"
			Set Rs=conn.Execute(sql)
		end if
		set rs1=nothing
		conn1.close
		set conn1=nothing
		if duidupd=0 then
			dim Allmjp
			Allmjp="1��|2��|3��|4��|5��|6��|7��|8��|9��|1Ͳ|2Ͳ|3Ͳ|4Ͳ|5Ͳ|6Ͳ|7Ͳ|8Ͳ|9Ͳ|1��|2��|3��|4��|5��|6��|7��|8��|9��|����|�Ϸ�|����|����|����|�װ�|����|"
			Allmjp=Allmjp & Allmjp & Allmjp & Allmjp
			mjpArr=split(Allmjp,"|")
			Randomize
			for i= 1 to 14
				mjpArr=split(Allmjp,"|")
				Mjx= Int(ubound(mjpArr)* Rnd)
				strMjx=mjpArr(Mjx)
				Allmjp=replace(Allmjp,strMjx & "|","",1,1,1)
				myMj=myMj & strMjx & "|"
			next
			Randomize
			for i= 1 to 13
				mjpArr=split(Allmjp,"|")
				Mjx= Int(ubound(mjpArr) * Rnd)
				strMjx=mjpArr(Mjx)
				Allmjp=replace(Allmjp,strMjx & "|","",1,1,1)
				youMj=youMj & strMjx & "|"
			next
			Allmjp_Akk=""
			Randomize
			for i= 1 to 109
				mjpArr = Split(Allmjp, "|")
				Mjx= Int(UBound(mjpArr) * Rnd)
				strMjx=mjpArr(Mjx)
				Allmjp=replace(Allmjp,strMjx & "|","",1,1,1)
				Allmjp_Akk=Allmjp_Akk & strMjx & "|"
			next
			sql="delete * from mjInfo where ׯ��='" & nickname & "' or ׯ��='" & fromname & "' or ����='" & nickname & "' or ����='" & fromname & "'"
			Set Rs=conn.Execute(sql)
			sql="insert into mjInfo(ׯ��,����,�齫) values ('"& nickname & "','"& fromname & "','"& Allmjp_Akk & "')"
			Set Rs=conn.Execute(sql)
			sql="select id from mjInfo where ׯ��='"& nickname & "'"
			Set Rs=conn.Execute(sql)
			mjID=Rs("id")
			sql="insert into dmj(ufrom,uto,Mymj,duz,isMy,isGet,isFp,mjID) values ('"& nickname & "','"& fromname & "','"& Mymj & "'," & xiazhu & ",true,false,true," & mjID & ")"
			Set Rs=conn.Execute(sql)
			sql="insert into dmj(ufrom,uto,Mymj,duz,isMy,isGet,isFp,mjID) values ('"& fromname & "','"& nickname & "','"& youMj & "'," & xiazhu & ",false,false,false," & mjID & ")"
			Set Rs=conn.Execute(sql)
			duidu="<b><font color=green>["&nickname&"]</font></b>ͬ����齫,<b><font color=green>["&application("aqjh_automanname")&"]</font></b>ϴ��......��������һ����<br>"
			duidu=duidu & "<b><font color=green>["&nickname&"]</font></b>����14����<img src='f2/mj/43.gif'>"
			duidu=duidu & "<input type=button value='ϴ��' onclick=""javascript:window.open('f2/dmj-xp.asp?name="&nickname&"','f3','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""dmjname""><img src='f2/mj/41.gif'>"
			duidu=duidu & " <b><font color=green>["
			duidu=duidu & fromname&"]</font></b>����13����<img src='f2/mj/43.gif'><input type=button value='ϴ��' onclick=""javascript:window.open('f2/dmj-xp.asp?name="
			duidu=duidu & fromname &"','f3','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""dmjname""><img src='f2/mj/41.gif'>"
			duidu=duidu & "<br>.....����[" & nickname & "]��ׯ�ң����ȴ��ơ�<br>"
			abc=1 
			msg="��ͬ����Է����齫�ˣ����Ժ�[������ʿ]���ڰ�����ϴ�ƣ�"
		end if
		set conn=nothing
		set rs=nothing
	end if
end if
if abc=1 then
says="<font color=#ff0000>�����齫��</font>��"&duidu			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
end if

%>
<html>
<head>
<style>
a:link {text-decoration:none; font-size:14;color:#000000}
a:hover {text-decoration:underline;font-size:14; color:#000000; background:#ffffff}
a:visited {text-decoration:none;font-size:14; color:#000000}
td {font-size:9pt; color:#ff0000; line-height:16pt}
</style>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>���齫</title>
</head>
<body bgcolor=#FFFFFF background="../bg.gif" >
<tr>
<td> <font color="red"> </font><br>
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#FFFFFF">
<tr> 
<td height="115"> 
          <table border="0" bgcolor="#0000FF" cellspacing="0" cellpadding="2" width="361">
            <tr> 
<td width="324" height="9"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">������ʾ</font><font color="red" size=2> 
</font></td>
<td width="29" height="9"> 
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr> 
<td width="16"><b><a href="<%if id="200" then%>javascript:window.close();<%else%>javascript:history.go(-1)<%end if%>" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="����"><font color="000000">��</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="359" cellpadding="4">
<tr> 
<td width="59" align="center" valign="top" height="29"><font face="Wingdings" color="#FF0000" style="font-size:32pt">J</font></td>
<td width="278" height="29"> <font color="red" size=2> 
<%=msg%>
</font></td>
</tr>
<tr> 
<td colspan="2" align="center" valign="top" height="58"> 
<input type=button value='�ر�' onClick="javascript:window.close()" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button223">
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
<%
function strMjp(cmj)
dim mj2
dim mjr
dim mjL
mjr=right(cmj,1)
if mjr=0 or cmj>40 then
mj2=replace(cmj,"10","����")
mj2=replace(cmj,"20","�Ϸ�")
mj2=replace(cmj,"30","����")
mj2=replace(cmj,"40","����")
mj2=replace(mj2,"41","����")
mj2=replace(mj2,"42","�װ�")
mj2=replace(mj2,"43","����")
else
mjL=Left(cmj,1)
mjL=replace(mjL,"1","��")
mj2=replace(mjL,"2","Ͳ")
mjL=replace(mjL,"3","��")
mj2=mjr & mjL
end if
strMjp=mj2
end function
function intMjp(cmj)
dim mj2
dim mjL
mj2=cmj
mjL left(cmj,1)
if isNumeric(mjL) then mj2=right(cmj,1) & mjL
mj2=replace(mj2,"��","1")
mj2=replace(mj2,"Ͳ","2")
mj2=replace(mj2,"��","3")
mj2=replace(mj2,"��","0")
mj2=replace(mj2,"��","1")
mj2=replace(mj2,"��","2")
mj2=replace(mj2,"��","3")
mj2=replace(mj2,"��","4")
mj2=replace(mj2,"����","41")
mj2=replace(mj2,"�װ�","42")
mj2=replace(mj2,"����","43")
intMjp=int(mj2)
end function
if to1="���" or to1=fromname then
%>
<script language="JavaScript">window.close();</script>
<%end if%>
</html>