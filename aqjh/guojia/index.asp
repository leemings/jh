<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>���һ���</title>
<LINK href="../css.css" rel=stylesheet>
</head><body leftmargin="0" bgcolor="#FFFFFF" topmargin="7" background=../bg.gif>
<p align=center><font size=3>��������ϵͳ</p>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr><td width=70 valign="top"></td>
    <td width="603" align="center" valign="top"> 
      <table width="100%" border="1" cellpadding="2" cellspacing="0" height="16">
        <tr align="center"> 
          <td width="79" height="11"><font color="#FF6600">����</font></td>
          <td width="66" height="11"><font color="#FF6600">����</font></td>
          <td width="40" height="11"><font color="#FF6600">������</font></td>
          <td width="40" height="11"><font color="#FF6600">ս����</font></td>
          <td width="80" height="11"><font color="#FF6600">���ұ���</font></td>
          <td width="103" height="11"><font color="#FF6600">����</font></td>
          <td width="152" height="11"><font color="#FF6600">���</font></td>
          <td width="55" height="11"><font color="#FF6600">�� ��</font></td>
        </tr>
        <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim jhmenpai(200,2)
rs.open "SELECT * FROM guo",conn
do while not rs.bof and not rs.eof
   	menpai=menpai+1
	jhmenpai(menpai,1)=trim(rs("a"))
	rs.movenext
loop
for xx=1 to menpai
tmprs=conn.execute("Select count(*) As ���� from �û� where ����='"&jhmenpai(xx,1)&"'")
regren=tmprs("����")
sql="update guo set c="& regren &" where a='"& jhmenpai(xx,1) &"'"
Set Rs=conn.Execute(sql)
next
rs.open "SELECT * FROM guo where a<>'��' and a<>'δ֪'",conn
do while not rs.bof and not rs.eof
%>
        <tr> 
          <td width="79" height="2"> 
            <div align="center"><a href="gj.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>"><%=rs("a")%></a></div>
          </td>
          <td width="66" height="2"> 
            <div align="center"><%=rs("b")%></div>
          </td>
          <td width="40" height="2"> 
            <div align="center"><%=rs("c")%></div>
          </td>          <td width="40" height="2"> 
            <div align="center"><%=rs("s")%></div>
          </td><td width="40" height="2"> 
            <div align="center"><%if rs("f")=1 then%>����<%else%>��<%end if%></div>
          </td>
          <td width="103" height="2"> 
            <div align="center">
          <%=rs("e")%>
          </div></td>
          <td width="152" height="2"> 
            <div align="center"><%=rs("d")%></div>
          </td>
          <td width="55" height="2"> 
            <div align="center"><a href="#" onClick="window.open('gj1.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')">����</a>|<a href="#" onClick="window.open('gj2.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')">�뿪</a></div>
          </td>
        </tr>
        <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
      </table>
    </td>
</tr>
</table>
<p align="center"><font color="#000000" size=2>��<font color="#FF0000"> �뿪���ң��ѹ��Ļ�,���Ǯ��������1000,��������С��!</font><br>
<font color="#0000FF">˭������һλ�õĹ��Ҿ�������,���������û���!</font></font>
</body>
</html>