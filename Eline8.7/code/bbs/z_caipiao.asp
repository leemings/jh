<!--#include file="CONN.ASP"-->
<!--#include file="z_conncaipiao.ASP"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="z_plus_check.asp"-->
<%
response.buffer=true
dim MinJiang,JiaGe,AMaxJiang1,AMaxJiang2,AMaxJiang3,AMaxJiang4,CP_Msg
dim CP_Qi,CP_MaxJiang,CP_TouZhuRenShu,CP_TouZhuE,CP_OpenDateTime,UserMoney
' ������С�������������ڴ������Զ�����(��Ҫ��һ�ڲſ���Ч)
MinJiang	=	20000
' ��ע�۸�
JiaGe		=	5
' ��ע��߽���
AMaxJiang1 =	15000
AMaxJiang2 =	1500
AMaxJiang3 =	150
AMaxJiang4 =	15


Stats="��Ʊ����"
call nav()
call head_var(3,0,"","")
response.write CP_Msg
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û�н����Ʊ���ĵ�Ȩ�ޣ�����<a href=login.asp><font color=blue>��¼</font></a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
set rs=server.createobject("adodb.recordset")

rs.open "select * from [user] where userid="&userid,conn,1,3
usermoney=rs("userwealth")
rs.close

dim rs3,rs1
set rs3=server.createobject("adodb.recordset")
rs3.open "select Top 1 * From CP_Record where OpenDateTime>=#"&year(now)&"-"&month(now)&"-"&day(now)&" "&hour(now)&":"&minute(now)&":"&second(now)&"# order by id desc",conncaipiao,1,3
if rs3.eof or rs3.bof then
	set rs1=server.createobject("adodb.recordset")
	rs1.open "select Top 1 * From CP_Record order by id desc",conncaipiao,1,3
	if rs1.eof or rs1.bof then
		CP_MaxJiang=MinJiang
	Else
		call KiaJiang(RS1("ID"))
		CP_MaxJiang=RS1("YuE")
	End if
	if CP_MaxJiang<MinJiang then
		CP_MaxJiang=CP_MaxJiang+MinJiang
	end if
	Rs1.close
	set rs1=nothing
	rs3.addnew
	Randomize
	rs3("Number")=int((8999*rnd)+1000)
	rs3("OpenDateTime")=date()+1&" 20:00:00"
	rs3("YuE")=CP_MaxJiang
	rs3.update
else
	set rs1=server.createobject("adodb.recordset")
	rs1.open "select Top 1 * From CP_user order by id desc",conncaipiao,1,3
	if not( rs1.bof and rs1.eof )then
		if rs1("Qi")<rs3("ID") then
			Randomize
			rs3("Number")=int((8999*rnd)+1000)
			rs3("OpenDateTime")=date()+1&" 20:00:00"
			rs3.update
		end if 
	end if			
end if
CP_Qi=rs3("id")
CP_OpenDateTime=rs3("OpenDateTime")
CP_MaxJiang=rs3("YuE")
rs3.close
set rs3=nothing
if request("job")<>"" then 
	call savepiao
	response.redirect scriptname
end if
CP_TouZhuRenShu=0
CP_TouZhuE=0

rs.open "select count(id) From CP_User Where Qi="&CP_Qi,conncaipiao,1,3
CP_TouZhuRenShu=rs(0)
rs.close
CP_MaxJiang=CP_MaxJiang+CP_TouZhuRenShu*JiaGe
CP_TouZhuE=CP_TouZhuRenShu*JiaGe
%>

<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
<tr><th><%=membername%>,��ӭ���ٲ�Ʊ����</th></tr>
<tr><td class=tablebody2>
<br>
<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style="word-break:break-all;">
<tr>
<TD vAlign=top class=tablebody1>
&nbsp;��ӭ���ٲ�Ʊ���ģ�<br>
&nbsp;�����ھͿ��Թ���� <font color=#FF0000><b><%=CP_Qi%></b></font> �ڲ�Ʊ�������ܹ��ۼƽ��� <font color=#FF0000><b><%=CP_MaxJiang%></b></font> Ԫ�� <br>
&nbsp;��ע��߽��� <font color=#0000FF><b>һ�Ƚ�</b></font>��<font color=#FF0000><b><%=AMaxJiang1%></b></font>Ԫ��<font color=#0000FF><b>���Ƚ�</b></font>��<font color=#FF0000><b><%=AMaxJiang2%></b></font>Ԫ��<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=#0000FF><b>���Ƚ�</b></font>��<font color=#FF0000><b><%=AMaxJiang3%></b></font>Ԫ��<font color=#0000FF><b>�ĵȽ�</b></font>��<font color=#FF0000><b><%=AMaxJiang4%></b></font>Ԫ��<br>
&nbsp;���ڹ���Ͷע <font color=#FF0000><b><%=CP_TouZhuRenShu%></b></font> ע����Ͷע�� <font color=#FF0000><b><%=CP_TouZhuE%></b></font> Ԫ���� <font color=#FF0000><b><%=CP_OpenDateTime%></b></font> ������<br>
</td>
<td vAlign=top class=tablebody1 rowspan=2>
<B>��ʷͶע:</b><br>
<%
	rs.open "select top 10 * from cp_user where user="&userid&" order by qi desc,id desc",conncaipiao,1,3
	i=0
	do while ( not rs.eof ) and i<10 
%>
��<%=rs("Qi")%>��:<%=right("0000"&rs("Haoma"),4)%><br>
<%
	 	rs.movenext
		i=i+1
	loop
	rs.close
%>
</td>
</tr>
<tr>
<TD vAlign=top class=tablebody1 align="center">
ÿע��Ʊ <font color=red><%=JiaGe%></font> Ԫ�������ڹ��� <font color=red><%=UserMoney%></font> Ԫ��<br>����Ҫ����ʲô���룿
<br><br>
<form action="z_CaiPiao.asp" method=Get name=BuyCaiPiao>
<input type="hidden" name="job" value="save">
<select name="CP1">
  <option selected>1</option>
  <option>2</option>
  <option>3</option>
  <option>4</option>
  <option>5</option>
  <option>6</option>
  <option>7</option>
  <option>8</option>
  <option>9</option>
  <option>0</option>
  <option>?</option>
</select>
<select name="CP2">
  <option selected>1</option>
  <option>2</option>
  <option>3</option>
  <option>4</option>
  <option>5</option>
  <option>6</option>
  <option>7</option>
  <option>8</option>
  <option>9</option>
  <option>0</option>
  <option>?</option>
</select>
<select name="CP3">
  <option selected>1</option>
  <option>2</option>
  <option>3</option>
  <option>4</option>
  <option>5</option>
  <option>6</option>
  <option>7</option>
  <option>8</option>
  <option>9</option>
  <option>0</option>
  <option>?</option>
</select>
<select name="CP4">
  <option selected>1</option>
  <option>2</option>
  <option>3</option>
  <option>4</option>
  <option>5</option>
  <option>6</option>
  <option>7</option>
  <option>8</option>
  <option>9</option>
  <option>0</option>
  <option>?</option>
</select>
<input type=submit value="Ͷע">
<input type="button" value="��ѡ" onclick="SelRandom()">
<script>
function SelRandom() {
	BuyCaiPiao.CP1.options[Math.random()*10].selected=true;
	BuyCaiPiao.CP2.options[Math.random()*10].selected=true;
	BuyCaiPiao.CP3.options[Math.random()*10].selected=true;
	BuyCaiPiao.CP4.options[Math.random()*10].selected=true;
}
SelRandom()
</script>
</form>
<%=CP_Msg%>
</td></tr>
</table>
<br>
</td></tr></table>
<br>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
<tr><th>��Ʊ����---����Ͷע���</th></tr>
<tr><td class=tablebody2>
<br>
	<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style="word-break:break-all;">
		<tr>
			<TD class=tablebody2 align=center height=20>�û���</td>
			<TD class=tablebody2 align=center>Ͷע��</td>
			<TD class=tablebody2 align=center>Ͷע���</td>  
		</tr>
<%
	rs.open "SELECT Count(UserName) AS UserNameOfCount,UserName FROM CP_User where qi="&CP_Qi&" GROUP BY UserName ORDER BY Count(UserName) DESC ",conncaipiao,1,3
	if rs.eof and rs.bof then
	%>
		<tr><td colspan=3 class=tablebody1 align=center height=20>���ڻ�û����Ͷע</td></tr>
	<%
	else
		do while not rs.eof
	%>		
			<tr>
				<TD class=tablebody1 align=center height=20><%=rs(1)%></td>
				<TD class=tablebody1 align=center><%=rs(0)%></td>
				<TD class=tablebody1 align=center><%=rs(0)*JiaGe%></td>
			</tr>
	<%
			rs.movenext
		loop
	end if
	rs.close
%>			
	</table>
<br>
</td></tr></table>
<br>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
<tr><th>��Ʊ����---���10���н����</th></tr>
<tr><td class=tablebody2>
<br>
	<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style="word-break:break-all;">
		<tr>
			<TD class=tablebody2 align=center height=20>��  ��</td>
			<TD class=tablebody2 align=center>�н�����</td>
			<TD class=tablebody2 align=center>��Ͷע��</td>
			<TD class=tablebody2 align=center>��Ͷע���</td>			
			<TD class=tablebody2 align=center>�����ܽ��</td>  
			<TD class=tablebody2 align=center>�н�ע��</td> 
			<TD class=tablebody2 align=center>�ҳ�����</td>
			<TD class=tablebody2 align=center>����ʣ�ཱ��</td> 			
		</tr>
<%
    set	rs=conncaipiao.execute("SELECT top 10 id,ZhongJiangHaoMa,ZongJiTouZhu,ZongJiTouZhujine,ZongJinE,ZhongJiangZhuShu,JiangJin,ShengY FROM qingkuang order by id DESC")
	if rs.eof and rs.bof then
	%>
		<tr><td colspan=8 class=tablebody1 align=center height=20>��û����ʷ�н����</td></tr>
	<%
	else
		do while not rs.eof
	%>		
			<tr>
			<TD class=tablebody1 align=center height=20><%=rs(0)%></td>
			<TD class=tablebody1 align=center><%=rs(1)%></td> 
			<TD class=tablebody1 align=center><%=rs(2)%></td>  
			<TD class=tablebody1 align=center><%=rs(3)%></td> 
			<TD class=tablebody1 align=center><%=rs(4)%></td> 
			<TD class=tablebody1 align=center><%=rs(5)%></td>
			<TD class=tablebody1 align=center><%=rs(6)%></td> 
			<TD class=tablebody1 align=center><%=rs(7)%></td> 
			</tr>
	<%
			rs.movenext
		loop
	end if
	rs.close
	set rs=nothing
%>			
	</table>
<br>
</td></tr></table>
<%
end if
conncaipiao.close
set conncaipiao=nothing
call activeonline()
call footer()

' ���������޸�
sub KiaJiang(ID)
	dim ZongJiTouZhu,ZongJinE,ZhongJiangHaoMa,ZhongJiangZhuShu1,ZhongJiangZhuShu2,ZhongJiangZhuShu3,ZhongJiangZhuShu4,JiangJin,Content,ShengY
	dim ZhongJiangUser,Delta  '�н�����   ���������
	' ����ܽ��
	rs.open "select * from CP_Record where ID="&id,conncaipiao,1,3
	ZongJinE=rs("YuE")
	ZhongJiangHaoMa=rs("Number")
	rs.close
	rs.open "select count(id) from CP_User where Qi="&ID,conncaipiao,1,3
	ZongJiTouZhu=rs(0)
	rs.close
	ZongJinE=ZongJinE+ZongJiTouZhu*JiaGe

	' ���һ�Ƚ�ע��
	rs.open "select count(ID) from CP_User where Qi="&id&" and HaoMa="&ZhongJiangHaoMa&"",conncaipiao,1,3    
	ZhongJiangZhuShu1=rs(0)  
	rs.close
	' ��ö��Ƚ�ע��
	rs.open "select count(ID) from CP_User where Qi="&id&" and HaoMa<>"&ZhongJiangHaoMa&" and (HaoMa mod 1000)="&(ZhongJiangHaoMa mod 1000)&"",conncaipiao,1,3    
	ZhongJiangZhuShu2=rs(0)
	rs.close
	' ������Ƚ�ע��
	rs.open "select count(ID) from CP_User where Qi="&id&" and (HaoMa mod 1000)<>"&int(ZhongJiangHaoMa mod 1000)&" and (HaoMa mod 100)="&int(ZhongJiangHaoMa mod 100)&"",conncaipiao,1,3
	ZhongJiangZhuShu3=rs(0)  
	rs.close
	' ����ĵȽ�ע��
	rs.open "select count(ID) from CP_User where Qi="&id&" and (HaoMa mod 100)<>"&(ZhongJiangHaoMa mod 100)&" and (HaoMa mod 10)="&(ZhongJiangHaoMa mod 10)&"",conncaipiao,1,3    
	ZhongJiangZhuShu4=rs(0)  
	rs.close
	' �ܽ���
	JiangJin=ZhongJiangZhuShu1*AMaxJiang1+ZhongJiangZhuShu2*AMaxJiang2+ZhongJiangZhuShu3*AMaxJiang3+ZhongJiangZhuShu4*AMaxJiang4
	' �ҽ���
	if JiangJin>ZongJinE then
		Delta=ZongJinE/JiangJin
	else
		Delta=1
	end if
	' ֧������
	JiangJin=ZhongJiangZhuShu1*int(AMaxJiang1*Delta)+ZhongJiangZhuShu2*int(AMaxJiang2*Delta)+ _
		ZhongJiangZhuShu3*int(AMaxJiang3*Delta)+ZhongJiangZhuShu4*int(AMaxJiang4*Delta)
	' ʣ�ཱ��
	ShengY=ZongJinE-JiangJin
	rs.open "select * from CP_Record where ID="&id,conncaipiao,1,3
	rs("YuE")=ShengY
	rs.update
	rs.close
	if ZhongJiangZhuShu1=0 and ZhongJiangZhuShu2=0 and ZhongJiangZhuShu3=0 and ZhongJiangZhuShu4=0 then
		ZhongJiangUser="���������н�"
	else
		ZhongJiangUser="�������н��ߡ�"
		ZhongJiangUser=ZhongJiangUser&chr(13)&"һ�Ƚ���"
		if ZhongJiangZhuShu1>0 then
			rs.open "select UserName,user,count(username),count(user) from CP_User where Qi="&id&" and HaoMa="&ZhongJiangHaoMa&" group by username,user" ,conncaipiao,1,3
			do while not rs.eof
				conn.execute("Update [user] set userWealth=userwealth+"&rs(2)*int(AMaxJiang1*Delta)&" where Userid="&rs(1))
				ZhongJiangUser=ZhongJiangUser&rs(0)&"*"&rs(2)&"  "
				rs.movenext
			loop
			rs.close
		else
			ZhongJiangUser=ZhongJiangUser&"��  "
		end if
		ZhongJiangUser=ZhongJiangUser&chr(13)&"���Ƚ���"
		if ZhongJiangZhuShu2>0 then
			rs.open "select UserName,user,count(username),count(user) from CP_User where Qi="&id&" and HaoMa<>"&ZhongJiangHaoMa&" and (HaoMa mod 1000)="&(ZhongJiangHaoMa mod 1000)&" group by username,user" ,conncaipiao,1,3
			do while not rs.eof
				conn.execute("Update [user] set userWealth=userwealth+"&rs(2)*int(AMaxJiang2*Delta)&" where Userid="&rs(1))
				ZhongJiangUser=ZhongJiangUser&rs(0)&"*"&rs(2)&"  "
				rs.movenext
			loop
			rs.close
		else
			ZhongJiangUser=ZhongJiangUser&"��  "
		end if
		ZhongJiangUser=ZhongJiangUser&chr(13)&"���Ƚ���"
		if ZhongJiangZhuShu3>0 then
			rs.open "select UserName,user,count(username),count(user) from CP_User where Qi="&id&" and (HaoMa mod 1000)<>"&(ZhongJiangHaoMa mod 1000)&" and (HaoMa mod 100)="&(ZhongJiangHaoMa mod 100)&" group by username,user" ,conncaipiao,1,3
			do while not rs.eof
				conn.execute("Update [user] set userWealth=userwealth+"&rs(2)*int(AMaxJiang3*Delta)&" where Userid="&rs(1))
				ZhongJiangUser=ZhongJiangUser&rs(0)&"*"&rs(2)&"  "
				rs.movenext
			loop
			rs.close
		else
			ZhongJiangUser=ZhongJiangUser&"��  "
		end if
		ZhongJiangUser=ZhongJiangUser&chr(13)&"�ĵȽ���"
		if ZhongJiangZhuShu4>0 then
			rs.open "select UserName,user,count(username),count(user) from CP_User where Qi="&id&" and (HaoMa mod 100)<>"&(ZhongJiangHaoMa mod 100)&" and (HaoMa mod 10)="&(ZhongJiangHaoMa mod 10)&" group by username,user" ,conncaipiao,1,3
			do while not rs.eof
				conn.execute("Update [user] set userWealth=userwealth+"&rs(2)*int(AMaxJiang4*Delta)&" where Userid="&rs(1))
				ZhongJiangUser=ZhongJiangUser&rs(0)&"*"&rs(2)&"  "
				rs.movenext
			loop
			rs.close
		else
			ZhongJiangUser=ZhongJiangUser&"��  "
		end if
	end if
	
	dim rs2
	set rs2=server.createobject("adodb.recordset")
	rs2.open "select username from cp_user where Qi="&id&" group by username",conncaipiao,1,3
	rs.open "select * from message where id is null",conn,3,3
	do while not rs2.eof
		rs.addnew
		rs("Sender")="��Ʊ����"
		rs("incept")=rs2("UserName")
		rs("Title")="��"&ID&"�ڲ�Ʊ��������"
		Content="�����н�����Ϊ"& ZhongJiangHaoMa &"���ܼ�Ͷע"&ZongJiTouZhu&"ע�������ܶ�"&ZongJinE&"Ԫ�������н�"& (ZhongJiangZhuShu1+ZhongJiangZhuShu2+ZhongJiangZhuShu3+ZhongJiangZhuShu4) &"ע,���ý���"& JiangJin &"Ԫ��"&chr(13)&"����ʣ�ཱ��"&ShengY&"Ԫ�������ڽ��ء�"
		Content=Content&" "&ZhongJiangUser
		rs("Content")=Content
		rs("issend")=1
		rs.update
		rs2.movenext
	loop
	rs.close
	rs2.close
	set rs2=nothing
	
	'д�н������  ������ 2002.10.16
	rs.open "select * from qingkuang where id is null",conncaipiao,3,3
	rs.addnew
	rs("ZhongJiangHaoMa")=ZhongJiangHaoMa
	rs("ZongJiTouZhu")=ZongJiTouZhu
	rs("ZongJiTouZhujine")=ZongJiTouZhu*JiaGe
	rs("ZongJinE")=ZongJinE
	rs("ZhongJiangZhuShu")=ZhongJiangZhuShu1+ZhongJiangZhuShu2+ZhongJiangZhuShu3+ZhongJiangZhuShu4
	rs("JiangJin")=JiangJin
	rs("ShengY")=ShengY
	rs.update
	rs.close	
end sub

sub savepiao
	dim cp1,cp2,cp3,cp4,xcount,i
	cp1=request("CP1")
	cp2=request("CP2")
	cp3=request("CP3")
	cp4=request("CP4")
	if cp1="" or cp2="" or cp3="" or cp4="" then
		CP_Msg="��������һ������û�й���������ѡ��"
	end if
	xcount=1
	if (cp1<"0" or cp1>"9") and cp1<>"?" then
		CP_Msg="��������һ������û�й���������ѡ��"
	else
		if cp1="?" then xcount=xcount*10
	end if
	if (cp2<"0" or cp2>"9") and cp2<>"?" then
		CP_Msg="��������һ������û�й���������ѡ��"
	else
		if cp2="?" then xcount=xcount*10
	end if
	if (cp3<"0" or cp3>"9") and cp3<>"?" then
		CP_Msg="��������һ������û�й���������ѡ��"
	else
		if cp3="?" then xcount=xcount*10
	end if
	if (cp4<"0" or cp4>"9") and cp4<>"?" then
		CP_Msg="��������һ������û�й���������ѡ��"
	else
		if cp4="?" then xcount=xcount*10
	end if
	if usermoney<JiaGe*xcount then
		CP_Msg="�Բ��������������Ҳ�������"&xcount&"ע��"
	end if
	if CP_Msg="" then
		rs.open "select * from CP_User where ID is null",conncaipiao,1,3
		for i=0 to 9999
			if (left(right("0000"&i,4),1)=cp1 or cp1="?") and _
				(mid(right("0000"&i,4),2,1)=cp2 or cp2="?") and _
				(mid(right("0000"&i,4),3,1)=cp3 or cp3="?") and _
				(right(right("0000"&i,4),1)=cp4 or cp4="?") then
				rs.addnew
				rs("Username")=Membername
				rs("User")=UserID
				rs("Haoma")=i
				rs("Qi")=CP_Qi
				rs.update
			end if
		next
		rs.close
		conn.execute("update [User] set userwealth=userwealth-"&JiaGe*xcount&" where userid="&userid)
		CP_Msg=cp1&cp2&cp3&cp4&"����ɹ��������Լ���ѡ����"
	end if
end sub
%>