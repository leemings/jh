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
id=trim(request.querystring("id"))
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
  duidu="���ܾ���["&nickname&"]�ܾ�["& fromname &"]�Ĵ��˿�������!"
  duidu=duidu & "<br>.........���˿��Ѹ�ý����һ�Ҫ�о�һ��..."
else
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	db="dpk.asp"
	'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
	'�����ķ�����֧��jet4.0����ʹ�����ص����ӷ�������߳�������
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.open connstr 
	sql="select * from dpk where ufrom='"& nickname & "' or ufrom='"& fromname & "'"
	Set Rs=conn.Execute(sql)
	if not (rs.eof and rs.bof) then
		unpkname=rs("ufrom")
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
		response.write "<script Language='Javascript'>alert('" & unpkname & "�����ƾ�û�н����������ٴο���')</script>"
		abc=1
		duidu="��������["&nickname&"]���ܽ���["& fromname &"]�ĶԶ�����!,��Ϊ["&unpkname&"]�����������ƾ�û�н�����"
		msg=unpkname & "�����ƾ�û�н����������ٴο��֣�"
	else
		sql="select duz from dpk where uto='"& nickname & "$��ע' and ufrom='"& fromname & "$��ע' order by id Desc"
		Set Rs=conn.Execute(sql)
		If rs.eof and rs.bof Then
			rs.close
			conn.close
			set rs=nothing
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('����û������ƾ֣�');}</script>"
			response.end
		end If
		xiazhuold=Rs(0)
		dim xia_fir,xiazhu,xia_sql,dppd,xia_class
		xia_fir=left(xiazhuold,1)
		xiazhu=mid(xiazhuold,2)
		xiazhu=clng(xiazhu)
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
				conn.close
				set rs=nothing
				set conn=nothing
				call ErrALT("�Ƿ�������")
		end select
		sql="delete * from dpk where uto='"& nickname & "$��ע' or ufrom='"& fromname & "$��ע' or ufrom='"& nickname & "$��ע' or uto='"& fromname & "$��ע'"
		Set Rs=conn.Execute(sql)
		dppd=0
		Set conn1=Server.CreateObject("ADODB.CONNECTION")
		Set rs1=Server.CreateObject("ADODB.RecordSet")
		conn1.open application("aqjh_usermdb")
		sql="select "& xia_sql &" from �û� where ����='"& nickname &"'"
		set rs1=conn1.execute(sql)
		sjxiazhu=rs1(0)
		rs1.close
		if xiazhu>sjxiazhu then
			msg=nickname&"�ĶĽ𲻹������ܿ��֡�"
			duidu="��������["&nickname&"]���ܽ���["& fromname &"]�ĶԶ�����!,��Ϊ["&nickname&"]�ĶĽ𲻹���"
			dppd=1
		end if
		sql="select "& xia_sql &" from �û� where ����='"& fromname &"'"
		set rs1=conn1.execute(sql)
		sjxiazhu=rs1(0)
		rs1.close
		if xiazhu>sjxiazhu then
			msg=fromname&"�ĶĽ𲻹������ܿ��֡�"
			duidu="��������["&nickname&"]���ܽ���["& fromname &"]�ĶԶ�����!,��Ϊ["&fromname&"]�ĶĽ𲻹���"
			dppd=1
		end if
		if dppd=0 then
				'����˫������ȥ�����¶�ע
			sql="update �û� set "& xia_sql &"="& xia_sql &"-"& xiazhu &" where ����='"& nickname &"' or ����='"& fromname &"'"
			set rs1=conn1.execute(sql)
		end if
		set rs1=nothing
		conn1.close
		set conn1=nothing
		abc=1
		if dppd=0 then
			for i= 1 to 18
				Randomize
				apk = Int(14 * Rnd+1)
				Randomize
				bbb= Int(4 * Rnd+1)
					if bbb=1 then bbb2="��"
				if bbb=2 then bbb2="��"
				if bbb=3 then bbb2="÷"
				if bbb=4 then bbb2="��"
				dpk=bbb2 & apk
				dpk=convpk(dpk)
				if instr(frpk,dpk & "|")=0 then
					frpk=frpk & dpk & "|"
				else 
					i=i-1
				end if
			next
			for i= 1 to 18
				Randomize
				apk = Int(14 * Rnd+1)
				Randomize
				bbb= Int(4 * Rnd+1)
				if bbb=1 then bbb2="��"
				if bbb=2 then bbb2="��"
				if bbb=3 then bbb2="÷"
				if bbb=4 then bbb2="��"
				dpk=bbb2 & apk
				dpk=convpk(dpk)
				if instr(frpk,dpk & "|")=0 and instr(topk,dpk & "|")=0 then
					topk=topk & dpk & "|"
				else 
					i=i-1
				end if
			next
			sql="insert into dpk(ufrom,uto,pk,fp,duz,oldpn) values ('"& nickname & "','"& fromname & "','"& topk & "',false," & xiazhuold & ",90)"
			Set Rs=conn.Execute(sql)
			sql="insert into dpk(ufrom,uto,pk,fp,duz,oldpn) values ('"& fromname & "','"& nickname & "','"& frpk & "',true," & xiazhuold & ",90)"
			Set Rs=conn.Execute(sql)
			duidu="<b><font color=green>["&nickname&"]</font></b>ͬ����˿�,<b><font color=green>["&application("aqjh_automanname")&"]</font></b>��ʼ����......<br>"
			duidu=duidu & "<b><font color=green>["&nickname&"]</font></b>ץ��18����<img src='f2/dpk/1.GIF'>"
			duidu=duidu & "<input type=button value='ϴ��' onclick=""javascript:window.open('f2/dpk-xp.asp?name="&nickname&"','f3','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""dpkname""><img src='f2/dpk/2.GIF'>"
			duidu=duidu & " <b><font color=green>["
			duidu=duidu & fromname&"]</font></b>ץ��18����<img src='f2/dpk/2.GIF'><input type=button value='ϴ��' onclick=""javascript:window.open('f2/dpk-xp.asp?name="
			duidu=duidu & fromname &"','f3','width=380,height=210')"" style=""background-color:#86A231;color:FFFFFF;border: 1 double"" onMouseOver=""this.style.color='FFFF00'"" onMouseOut=""this.style.color='FFFFFF'"" name=""dpkname""><img src='f2/dpk/1.GIF'>"
			duidu=duidu & "<br>.....�����ƾֹ涨����[" & nickname & "]�ȷ��ơ�<br>"
			abc=1
			msg="��ͬ����Է����˿��ˣ����Ժ�[������ʿ]���ڰ�����ϴ�ƣ�"
			conn.close
			set rs=nothing
			set conn=nothing
		end if
	end if
end if

if abc=1 then
says="<font color=#ff0000>�����˿ˡ�</font>��"&duidu			'��������
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
<title>�˺㽭��-���˿�</title>
</head>
<body bgcolor=#FFFFFF background="../bg.gif" >
<tr>
<td> <font color="red"> </font><br>
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#FFFFFF">
<tr> 
<td height="115"> 
          <table border="0" bgcolor="#0000FF" cellspacing="0" cellpadding="2" width="361">
            <tr> 
<td width="324" height="9"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">�˺㽭����ʾ</font><font color="red" size=2> 
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
function convpk(dpk)
if dpk="��14" then dpk="����"
if dpk="��14" then dpk="����"
if dpk="÷14" then dpk="С��"
if dpk="��14" then dpk="С��"
dpk=replace(dpk,"13","K")
dpk=replace(dpk,"12","Q")
dpk=replace(dpk,"11","J")
dpk=replace(dpk,"10","ʮ")
dpk=replace(dpk,"1","A")
dpk=replace(dpk,"ʮ","10")
convpk=dpk
end function 
if to1=jiutian_dabusi then
%>
<script language="JavaScript">window.close();</script>
<%end if%>
</html>

