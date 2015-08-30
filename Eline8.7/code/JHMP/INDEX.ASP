<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>江湖门派♀wWw.happyjh.com♀</title>
<LINK href="../css.css" rel=stylesheet>
</head><body
leftmargin="0" bgcolor="#FFFFFF" topmargin="7">
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="175" valign="top"><img src="../images/mp1.gif"><img src="../images/mp2.gif"><img src="../images/mp3.gif"></td>
    <td width="603" align="center" valign="top"> 
      <table width="100%" border="1" cellpadding="2" cellspacing="0" height="16">
        <tr align="center"> 
          <td width="79" height="11"><font color="#FF6600">门派</font></td>
          <td width="66" height="11"><font color="#FF6600">掌门</font></td>
          <td width="40" height="11"><font color="#FF6600">弟子数</font></td>
          <td width="173" height="11"><font color="#FF6600">限制</font></td>
          <td width="152" height="11"><font color="#FF6600">简介</font></td>
          <td width="55" height="11"><font color="#FF6600">操 作</font></td>
        </tr>
        <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim jhmenpai(200,2)
rs.open "SELECT * FROM p",conn
do while not rs.bof and not rs.eof
   	menpai=menpai+1
	jhmenpai(menpai,1)=trim(rs("a"))
	rs.movenext
loop
for xx=1 to menpai
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 门派='"&jhmenpai(xx,1)&"'")
regren=tmprs("数量")
sql="update p set c="& regren &" where a='"& jhmenpai(xx,1) &"'"
Set Rs=conn.Execute(sql)
next
rs.open "SELECT * FROM p where a<>'官府' and a<>'游侠'",conn
do while not rs.bof and not rs.eof
%>
        <tr> 
          <td width="79" height="2"> 
            <div align="center"><a href="mp.asp?my=<%=sjjh_name%>&amp;id=<%=rs("a")%>"><%=rs("a")%></a></div>
          </td>
          <td width="66" height="2"> 
            <div align="center"><%=rs("b")%></div>
          </td>
          <td width="40" height="2"> 
            <div align="center"><%=rs("c")%></div>
          </td>
          <td width="173" height="2"> 
            <div align="center">
          <%=rs("e")%>
          </div></td>
          <td width="152" height="2"> 
            <div align="center"><%=rs("d")%></div>
          </td>
          <td width="55" height="2"> 
            <div align="center"><a href="#" onClick="window.open('mp1.asp?my=<%=sjjh_name%>&amp;id=<%=rs("a")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')">加入</a>|<a href="#" onClick="window.open('mp2.asp?my=<%=sjjh_name%>&amp;id=<%=rs("a")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')">离开</a></div>
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
<p align="center"><font color="#000000">　<font color="#FF0000"> 离开门派判帮的话,你的钱、存款将清至1000,所以请你小心!</font><br>
<font color="#0000FF">谁都想找一位好的老大照着,加入门派是没错的!</font></font>
</body>
</html>
