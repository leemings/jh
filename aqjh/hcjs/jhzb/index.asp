<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
function myzbzt(newnai,nai)
	ztz=int((newnai/nai)*100)
	if ztz>=95 then
		myzbzt="<font color=green>"&ztz&"��</font>  ���"
	else
		if ztz>=75 and ztz<95 then
			myzbzt="<font color=#004080>"&ztz&"��</font>  ��΢��"
		else
			if ztz>=55 and ztz<75 then
				myzbzt="<font color=#ff00ff>"&ztz&"��</font>  ��"
			else
				if ztz>=35 and ztz<55 then
					myzbzt="<font color=red>"&ztz&"��</font>  ������"
				else
					if ztz>=15 and ztz<35 then
						myzbzt="<font color=blue>"&ztz&"��</font>  ��������"
					else
						myzbzt=ztz&"��  ����(�޷��޸�)"
					end if
				end if
			end if
		end if
	end if
end function

function whsm(nei,newnei,lb,whcs)
	if isnull(rs(lb)) or trim(rs(lb))="" then
		whsm="û��װ��"
		exit function
	end if
	if whcs>=3 then
		whsm="����ά�� <a href='pqzb.asp?wqlb="&lb&"'><font color=blue>����</font></a>"
		exit function
	end if
	if newnai<>0 and nai<>0 then
		ztz=int((newnai/nai)*100)
	else
		ztz=0
	end if
	if ztz<15 then
		whsm="�޷��޸� <a href='pqzb.asp?wqlb="&lb&"'><font color=blue>����</font></a>"
		exit function
	end if
	select case whcs
		case 0
			if ztz<95 then
				whsm="<a href='wxwq.asp?lb="&lb&"'>ά��</a> <a href='pqzb.asp?wqlb="&lb&"'><font color=blue>����</font></a>"
			else
				whsm="����ά�� <a href='pqzb.asp?wqlb="&lb&"'><font color=blue>����</font></a>"
			end if
		case 1
			if ztz<75 then
				whsm="<a href='wxwq.asp?lb="&lb&"'>ά��</a> <a href='pqzb.asp?wqlb="&lb&"'><font color=blue>����</font></a>"
			else
				whsm="����ά�� <a href='pqzb.asp?wqlb="&lb&"'><font color=blue>����</font></a>"
			end if
		case 2
			if ztz<55 then
				whsm="<a href='wxwq.asp?lb="&lb&"'>ά��</a> <a href='pqzb.asp?wqlb="&lb&"'><font color=blue>����</font></a>"
			else
				whsm="����ά�� <a href='pqzb.asp?wqlb="&lb&"'><font color=blue>����</font></a>"
			end if
	end select
end function

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� WHERE ����='"&aqjh_name&"'",conn,3,3
tlj=rs("������")
nlj=rs("������")
%>
<head>
<title>�ҵ�װ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body background="BACK.GIF" oncontextmenu=self.event.returnValue=false>
<p align="center">
<table width="620" border="0" cellspacing="0" cellpadding="0" height="324">
  <tr> 
    <td rowspan="8" height="324" width="71" valign="top"> 
       <font size="2">
	   <input onClick="javascript:window.document.location.href='index.asp'" title="װ��һ��" type=button value="װ��һ��" name="button1">
      <br>
      <input onClick="javascript:window.document.location.href='daojun.asp'" title="����װ��" type=button value="����װ��" name="button">
       <br>
       <br>
<a href="#" onClick="window.open('wxsm.htm','wxsm','scrollbars=no,resizable=no,width=690,height=440')">����鿴װ��ά��˵��</a></font></td>
    <td colspan="9" height="34"> 
      <div align="center"><font size="2">�û�<font color="#FF0000"><%=rs("����")%></font>״̬����ɫΪ���ޣ�<font color="#000000">���ӣ�<%=rs("����")%> 
        ���飺<%=rs("allvalue")%> �ȼ���<%=rs("�ȼ�")%>��</font><br>
        ��������<%=rs("����")%> ��������<%=rs("����")%> ������<%=rs("����")%><font color="#FF0000">/<%=rs("�ȼ�")*64+2000%></font>   
        ������<%=rs("����")%><font color="#FF0000">/<%=rs("�ȼ�")*240+5260%></font></font></div>  
    </td>
  </tr>
  <tr> 
    <td width="40" height="36" bordercolor="#000000" style="border: 1 solid #000000" align="center"><font size="2">���</font></td>
    <td width="137" height="36" bordercolor="#000000" style="border: 1 solid #000000" align="center"><font size="2">װ����</font></td>
    <td width="51" height="36" bordercolor="#000000" style="border: 1 solid #000000" align="center"><font size="2">����</font></td>
    <td width="52" height="36" bordercolor="#000000" style="border: 1 solid #000000" align="center"><font size="2">����</font></td>
    <td width="41" height="36" bordercolor="#000000" style="border: 1 solid #000000" align="center"><font size="2">����</font></td>
    <td width="68" height="36" bordercolor="#000000" style="border: 1 solid #000000" align="center"><font size="2">�;ö�</font></td>
    <td width="92" height="36" bordercolor="#000000" style="border: 1 solid #000000" align="center"><font size="2">�𻵳̶�</font></td>
    <td width="35" height="36" bordercolor="#000000" style="border: 1 solid #000000" align="center"><font size="2">ά��</font></td>
    <td width="33" height="36" bordercolor="#000000" style="border: 1 solid #000000" align="center"><font size="2">����</font></td>
  </tr>
  <tr> 
    <td width="40" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2">ͷ��</font></td>
    <%if isnull(rs("z1")) or trim(rs("z1"))="" then
			myzb="δװ���κ���Ʒ��"
			gj=0
			fy=0
			nai=0
			newnai=0
			whcs=0
			zbzt="��"
			tx="��"
		else
			data1=split(rs("z1"),"|")
			sql="select * FROM b WHERE a='" & data1(0) &"'"
			set rs1=conn.execute(sql)
			If Rs1.Bof OR Rs1.Eof then
				myzb="δװ���κ���Ʒ��"
				gj=0
				fy=0
				nai=0
				newnai=0
				whcs=0
				zbzt="��"
				tx="��"
			else
				myzb="<img src='../jhjs/images/"&rs1("i")&"' alt='"&rs1("c")&"'>["&rs1("a")&"]"
				gj=rs1("f")
				fy=rs1("g")
				nai=rs1("j")
				newnai=data1(1)
				tx=rs1("k")
				whcs=data1(5)
				zbzt=myzbzt(newnai,nai)
			end if
			rs1.close
		end if%>
    <td width="137" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=myzb%></font></td>
    <td width="51" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#804040"><%=gj%></font></td>
    <td width="52" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#804040"><%=fy%></font></td>
    <td width="41" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#8000FF"><%=tx%></font></td>
    <td width="68" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><font color=red><%=newnai%></font>/<font color=green><%=nai%></font></font></td>
    <td width="92" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=zbzt%></font></td>
    <td width="35" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#FF0000"><%=whcs%></font><font size="2">��</font></td>
    <td width="33" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=whsm(nei,newnei,"z1",whcs)%></font></td>
  </tr>
  <tr> 
    <td width="40" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2">װ��</font></td>
    <%if isnull(rs("z2")) or trim(rs("z2"))="" then
			myzb="δװ���κ���Ʒ��"
			gj=0
			fy=0
			nai=0
			newnai=0
			whcs=0
			zbzt="��"
			tx="��"
		else
			data1=split(rs("z2"),"|")
			sql="select * FROM b WHERE a='" & data1(0) &"'"
			set rs1=conn.execute(sql)
			If Rs1.Bof OR Rs1.Eof then
				myzb="δװ���κ���Ʒ��"
				gj=0
				fy=0
				nai=0
				newnai=0
				whcs=0
				zbzt="��"
				tx="��"
			else
				myzb="<img src='../jhjs/images/"&rs1("i")&"' alt='"&rs1("c")&"'>["&rs1("a")&"]"
				gj=rs1("f")
				fy=rs1("g")
				nai=rs1("j")
				newnai=data1(1)
				tx=rs1("k")
				whcs=data1(5)
				zbzt=myzbzt(newnai,nai)
			end if
			rs1.close
		end if%>
    <td width="137" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=myzb%></font></td>
    <td width="51" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#804040"><%=gj%></font></td>
    <td width="52" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#804040"><%=fy%></font></td>
    <td width="41" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#8000FF"><%=tx%></font></td>
    <td width="68" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><font color=red><%=newnai%></font>/<font color=green><%=nai%></font></font></td>
    <td width="92" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=zbzt%></font></td>
    <td width="35" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#FF0000"><%=whcs%></font><font size="2">��</font></td>
    <td width="33" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=whsm(nai,newnai,"z2",whcs)%></font></td>
  </tr>
  <tr> 
    <td width="40" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2">����</font></td>
    <%if isnull(rs("z3")) or trim(rs("z3"))="" then
			myzb="δװ���κ���Ʒ��"
			gj=0
			fy=0
			nai=0
			newnai=0
			whcs=0
			zbzt="��"
			tx="��"
		else
			data1=split(rs("z3"),"|")
			sql="select * FROM b WHERE a='" & data1(0) &"'"
			set rs1=conn.execute(sql)
			If Rs1.Bof OR Rs1.Eof then
				myzb="δװ���κ���Ʒ��"
				gj=0
				fy=0
				nai=0
				newnai=0
				whcs=0
				zbzt="��"
				tx="��"
			else
				myzb="<img src='../jhjs/images/"&rs1("i")&"' alt='"&rs1("c")&"'>["&rs1("a")&"]"
				gj=rs1("f")
				fy=rs1("g")
				nai=rs1("j")
				newnai=data1(1)
				tx=rs1("k")
				whcs=data1(5)
				zbzt=myzbzt(newnai,nai)
			end if
			rs1.close
		end if%>
    <td width="137" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=myzb%></font></td>
    <td width="51" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#804040"><%=gj%></font></td>
    <td width="52" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#804040"><%=fy%></font></td>
    <td width="41" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#8000FF"><%=tx%></font></td>
    <td width="68" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><font color=red><%=newnai%></font>/<font color=green><%=nai%></font></font></td>
    <td width="92" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=zbzt%></font></td>
    <td width="35" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#FF0000"><%=whcs%></font><font size="2">��</font></td>
    <td width="33" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=whsm(nai,newnai,"z3",whcs)%></font></td>
  </tr>
  <tr> 
    <td width="40" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2">����</font></td>
    <%if isnull(rs("z4")) or trim(rs("z4"))="" then
			myzb="δװ���κ���Ʒ��"
			gj=0
			fy=0
			nai=0
			newnai=0
			whcs=0
			zbzt="��"
			tx="��"
		else
			data1=split(rs("z4"),"|")
			sql="select * FROM b WHERE a='" & data1(0) &"'"
			set rs1=conn.execute(sql)
			If Rs1.Bof OR Rs1.Eof then
				myzb="δװ���κ���Ʒ��"
				gj=0
				fy=0
				nai=0
				newnai=0
				whcs=0
				zbzt="��"
				tx="��"
			else
				myzb="<img src='../jhjs/images/"&rs1("i")&"' alt='"&rs1("c")&"'>["&rs1("a")&"]"
				gj=rs1("f")
				fy=rs1("g")
				nai=rs1("j")
				newnai=data1(1)
				tx=rs1("k")
				whcs=data1(5)
				zbzt=myzbzt(newnai,nai)
			end if
			rs1.close
		end if%>
    <td width="137" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=myzb%></font></td>
    <td width="51" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#804040"><%=gj%></font></td>
    <td width="52" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#804040"><%=fy%></font></td>
    <td width="41" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#8000FF"><%=tx%></font></td>
    <td width="68" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><font color=red><%=newnai%></font>/<font color=green><%=nai%></font></font></td>
    <td width="92" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=zbzt%></font></td>
    <td width="35" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2" color="#FF0000"><%=whcs%></font><font size="2">��</font></td>
    <td width="33" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="42"><font size="2"><%=whsm(nai,newnai,"z4",whcs)%></font></td>
  </tr>
  <tr> 
    <td width="40" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2">����</font></td>
    <%if isnull(rs("z5")) or trim(rs("z5"))="" then
			myzb="δװ���κ���Ʒ��"
			gj=0
			fy=0
			nai=0
			newnai=0
			whcs=0
			zbzt="��"
			tx="��"
		else
			data1=split(rs("z5"),"|")
			sql="select * FROM b WHERE a='" & data1(0) &"'"
			set rs1=conn.execute(sql)
			If Rs1.Bof OR Rs1.Eof then
				myzb="δװ���κ���Ʒ��"
				gj=0
				fy=0
				nai=0
				newnai=0
				whcs=0
				zbzt="��"
				tx="��"
			else
				myzb="<img src='../jhjs/images/"&rs1("i")&"' alt='"&rs1("c")&"'>["&rs1("a")&"]"
				gj=rs1("f")
				fy=rs1("g")
				nai=rs1("j")
				newnai=data1(1)
				tx=rs1("k")
				whcs=data1(5)
				zbzt=myzbzt(newnai,nai)
			end if
			rs1.close
		end if%>
    <td width="137" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2"><%=myzb%></font></td>
    <td width="51" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2" color="#804040"><%=gj%></font></td>
    <td width="52" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2" color="#804040"><%=fy%></font></td>
    <td width="41" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2" color="#8000FF"><%=tx%></font></td>
    <td width="68" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2"><font color=red><%=newnai%></font>/<font color=green><%=nai%></font></font></td>
    <td width="92" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2"><%=zbzt%></font></td>
    <td width="35" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2" color="#FF0000"><%=whcs%></font><font size="2">��</font></td>
    <td width="33" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2"><%=whsm(nai,newnai,"z5",whcs)%></font></td>
  </tr>
  <tr> 
    <td width="40" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2">˫��</font></td>
    <%if isnull(rs("z6")) or trim(rs("z6"))="" then
			myzb="δװ���κ���Ʒ��"
			gj=0
			fy=0
			nai=0
			newnai=0
			whcs=0
			zbzt="��"
			tx="��"
		else
			data1=split(rs("z6"),"|")
			sql="select * FROM b WHERE a='" & data1(0) &"'"
			set rs1=conn.execute(sql)
			If Rs1.Bof OR Rs1.Eof then
				myzb="δװ���κ���Ʒ��"
				gj=0
				fy=0
				nai=0
				newnai=0
				whcs=0
				zbzt="��"
				tx="��"
			else
				myzb="<img src='../jhjs/images/"&rs1("i")&"' alt='"&rs1("c")&"'>["&rs1("a")&"]"
				gj=rs1("f")
				fy=rs1("g")
				nai=rs1("j")
				newnai=data1(1)
				tx=rs1("k")
				whcs=data1(5)
				zbzt=myzbzt(newnai,nai)
			end if
			rs1.close
			set rs1=nothing
		end if%>
    <td width="137" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2"><%=myzb%></font></td>
    <td width="51" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2" color="#804040"><%=gj%></font></td>
    <td width="52" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2" color="#804040"><%=fy%></font></td>
    <td width="41" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2" color="#8000FF"><%=tx%></font></td>
    <td width="68" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2"><font color=red><%=newnai%></font>/<font color=green><%=nai%></font></font></td>
    <td width="92" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2"><%=zbzt%></font></td>
    <td width="35" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2" color="#FF0000"><%=whcs%></font><font size="2">��</font></td>
    <td width="33" bordercolor="#000000" style="border: 1 solid #000000" align="center" height="43"><font size="2"><%=whsm(nai,newnai,"z6",whcs)%></font></td>
  </tr>
</table><br><center><FONT 
            color=#0000ff>&copy; ��Ȩ���� 2005-2006 </FONT><A 
            href="http://www.7758530.com/" target=_blank><FONT 
            color=#0000ff>���ֽ�����</FONT></A>
<%rs.close
set rs=nothing
conn.close
set conn=nothing
%>