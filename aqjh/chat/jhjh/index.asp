<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<%
sub gfsj(gfdata)
if not(isnull(gfdata)) and instr(gfdata,"|")<>0 then
	mydata=split(gfdata,"|")
	mygj=mygj+mydata(2)
	myfy=myfy+mydata(3)
	erase mydata
end if
end sub
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<!��IP>
<%
ip=request("ip")
if ip="" then
ip=Request.ServerVariables("REMOTE_ADDR")
end if
sip=split(ip,".")
if ubound(sip)<>3 then
ip=Request.ServerVariables("REMOTE_ADDR")
erase sip
sip=split(ip,".")
end if
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
rs.open "select * FROM z WHERE a<=" & num & " and b>=" & num,conn,1,1
if rs.eof or rs.bof then
country="����"
city="δ֪"
else
country=rs("c")
city=rs("d")
end if
if country="" then country="�й�"
if city="" then city="δ֪"
last="����:" & country & " ����:" & city
rs.close
%>
<!��IP>
<%
rs.open "Select * from �û� where ����='" & aqjh_name &"'",conn,3,3
gjsx=rs("�ȼ�")*aqjh_gjsx
fysx=rs("�ȼ�")*aqjh_fysx
if rs("��Ա")=true and clng(DateDiff("d",now(),rs("��Ա����")))>0 then
	pdstr="<font size=-1>[�ݵ��ƻ�Ա]"&rs("��Ա����")&"</font>"
else
	pdstr=""
end if
wgj=rs("�书��")
nlj=rs("������")
tlj=rs("������")
jhjh=rs("����")
mygj=rs("����")
myfy=rs("����")
mywg=rs("�书")
call gfsj(rs("z1"))
call gfsj(rs("z2"))
call gfsj(rs("z3"))
call gfsj(rs("z4"))
call gfsj(rs("z5"))
call gfsj(rs("z6"))
if rs("����1")<>"����" and rs("����1")<>"" then
	zbdata=split(rs("����1"),"|")
	xs=0
	if UBound(zbdata)=5 then
		sj=datediff("n",zbdata(1),now())
		jgj=zbdata(2)
		jfy=zbdata(3)
		jwg=zbdata(4)
		if sj<=60 then
			mygj=mygj+zbdata(2)
			myfy=myfy+zbdata(3)
			mywg=mywg+zbdata(4)
			xs=1
		end if
	end if
	erase zbdata
end if
%><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<HTML>
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {
	color: #FF0000;
	font-weight: bold;
}
.style4 {color: #0000FF}
-->
</style>
<HEAD><TITLE>���ֽ���ʱ������ϵͳ</TITLE>
<body bgcolor="#739EEF" background="">

<META content=zh-cn http-equiv=Content-Language>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<STYLE>
.word {
	COLOR: #333333; FONT-FAMILY: "����"; FONT-SIZE: 9pt; LETTER-SPACING: 0.5pt; LINE-HEIGHT: 17pt; WORD-SPACING: 2pt
}
A:link {
	COLOR: #000000; FONT-STYLE: normal; TEXT-DECORATION: none
}
A:visited {
	COLOR: #000000; FONT-STYLE: normal; TEXT-DECORATION: none
}
A:hover {
	COLOR: #0066cc; FONT-STYLE: normal; TEXT-DECORATION: none
}
A:active {
	COLOR: #0066cc; FONT-STYLE: normal; TEXT-DECORATION: none
}
.en {
	FONT-FAMILY: "Arial", "Helvetica", "sans-serif"; FONT-SIZE: 12px
}
.style1 {
	color: #FF0000;
	font-weight: bold;
}
</STYLE>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" width="100%">
  <TBODY>
  <TR align=middle vAlign=center>
    <TD>
      <TABLE 
      width=690 height="382" border=0 cellPadding=0 cellSpacing=0 bgColor=#ffffff 
      style="BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid">
        <TBODY>
        <TR bgColor=#666666>
          <TD colSpan=2 height=3></TD></TR>
        <TR>
          <TD width=31 vAlign=top bgcolor="#DEE7FF">&nbsp;</TD>
          <TD width="657" vAlign=top bgcolor="#DEE7FF">
            <TABLE border=1 cellPadding=0 cellSpacing=0 width="100%">
              <TBODY>
              <TR bordercolor="#FFFFFF" bgcolor="#DEE7FF">
                <td height="27" colspan="6" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 10pt; font-weight: bold; color: #CC3399;">��ӭ���ٿ��ֽ���ʱ��������<a href="help.htm" target="_blank">����鿴��������ϵͳ</a></span>
                  <p align="center"><font size=2 color=#3366CC><%=aqjh_name%></FONT><span class="style1"><font size=2>�����ڵĽ���������:</FONT></span><font size=2></FONT><font size=2 color=#3366CC><%=rs("����")%></FONT></p></td>
                </TR>
			  <TR bordercolor="#FFFFFF" bgcolor="#DEE7FF">
                <td width="85" height="34" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="color: #FF0000; font-weight: bold;"><font style="font-size: 9pt">��������</font></span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-weight: bold; color: #FF0000;"><font style="font-size: 9pt">��ʾͼ��</font></span></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-weight: bold; color: #FF0000;"><font style="font-size: 9pt">������Ҫ����</font></span></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span class="style2">���Ľ��</span></td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-weight: bold; color: #FF0000;"><font style="font-size: 9pt">�õ�����</font></span></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-weight: bold; color: #FF0000;"><font style="font-size: 9pt">���Ͻ���</font></span></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">��ѧ����</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">��ͼ����ʾ</span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��&gt;1���ߵȼ�&gt;20</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">100</span></td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+10</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><span style="font-size: 9pt"><a href="1.asp">����=>ԭʼ��</a></span></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">ԭʼ��</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_star.gif" width="13" height="12"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��&gt;1���ߵȼ�&gt;50</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">500</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+20</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="2.asp">����=>�Ҷ���</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">�Ҷ���</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_star.gif" width="13" height="12"><img src="jhjh_files/time_star.gif" width="13" height="12"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��&gt;1���ߵȼ�&gt;100</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">1000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+30</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="3.asp">����=>������</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">������</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_star.gif" width="13" height="12"><img src="jhjh_files/time_star.gif" width="13" height="12"><img src="jhjh_files/time_star.gif" width="13" height="12"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��&gt;2���ߵȼ�&gt;300</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">1500</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+40</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="4.asp">����=&gt;ľͷ��</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">ľͷ��</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_yueliang.gif" width="15" height="15"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��&gt;2���ߵȼ�&gt;400</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">6000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+50</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="5.asp">����=&gt;С����</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">С����</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><strong style="font-weight: 400"><span style="font-size: 9pt"><img src="jhjh_files/time_yueliang.gif" width="15" height="15"><img src="jhjh_files/time_star.gif" width="13" height="12"></span></strong></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��&gt;2���ߵȼ�&gt;700</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">9000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+60</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="6.asp">����=&gt;�׸���</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">�׸���</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_yueliang.gif" width="15" height="15"><img src="jhjh_files/time_star.gif" width="13" height="12"><img src="jhjh_files/time_star.gif" width="13" height="12"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��3</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">15000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+70</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="7.asp">����=&gt;�ִ���</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt"><span style="font-size: 9pt">�ִ���</span></span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_yueliang.gif" width="15" height="15"><img src="jhjh_files/time_star.gif" width="13" height="12"><img src="jhjh_files/time_star.gif" width="13" height="12"><img src="jhjh_files/time_star.gif" width="13" height="12"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��4</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">20000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+80</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="8.asp">����=&gt;δ����</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt"><span style="font-size: 9pt">δ����</span></span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_yueliang.gif" width="15" height="15"><img src="jhjh_files/time_yueliang.gif" width="15" height="15"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��5</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">30000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+90</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="9.asp">����=&gt;̫����</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">̫����</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_yueliang.gif" width="15" height="15"><img src="jhjh_files/time_yueliang.gif" width="15" height="15"><img src="jhjh_files/time_yueliang.gif" width="15" height="15"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��6</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">50000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+100</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="10.asp">����=&gt;������</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">������</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_sun.gif" width="16" height="16" border="0"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��7</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">60000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+110</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="11.asp">����=&gt;��Ӱ��</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">��Ӱ��</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_sun.gif" width="16" height="16"><img src="jhjh_files/time_sun.gif" width="16" height="16" border="0"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��8</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">80000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+120</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="12.asp">����=&gt;������</a></span></strong></td>
              </TR>
			  <TR bordercolor="#66CCFF" bgcolor="#DEE7FF">
                <td width="85" height="25" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><span style="font-size: 9pt">������</span></td>
                <td width="110" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt"><img src="jhjh_files/time_sun.gif" width="16" height="16"><img src="jhjh_files/time_sun.gif" width="16" height="16" border="0"><img src="jhjh_files/time_sun.gif" width="16" height="16" border="0"></span></font></td>
                <td width="99" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><font style="font-size: 9pt"><span style="font-size: 9pt">ת��9</span></font></td>
                <td width="114" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">200000</td>
                <td width="130" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt">������+130</span></strong></td>
                <td width="105" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF"><strong style="font-weight: 400"><span style="font-size: 9pt"><a href="13.asp">����=&gt;ȫ����</a></span></strong></td>
              </TR>
              </TBODY></TABLE>            </TD>
        </TR>
		<tr bgcolor="#FFFFFF"> 
          <td bgcolor="#CCCCFF" colspan="3"><div align="center"><span class="style4">&copy; ��Ȩ���� 2015-2015 <A href="http://www.happyjh.com/" target=_blank>���ֽ�����վ</A></span> </div></td>
        </tr>
			  
        <TR bgColor=#686468>
        <TD colSpan=2 
height=4></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></BODY></HTML>
