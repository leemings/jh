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
<title><%=Application("aqjh_chatroomname")%>江湖门派排行</title>
<LINK href="../css.css" rel=stylesheet>
</head><body leftmargin="0" bgcolor="#FFFFFF" topmargin="7" background=../bg.gif>
<p align=center><font size=3>快乐江湖门派属性排行</p>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr><td width=70 valign="top"></td>
    <td width="603" align="center" valign="top"> 
      <table width="100%" border="1" cellpadding="2" cellspacing="0" height="16">
        <tr align="center"> 
          <td width="79" height="11"><font color="#FF6600">门派</font></td>
          <td width="66" height="11"><font color="#FF6600">等级</font></td>
          <td width="40" height="11"><font color="#FF6600">金币</font></td>
          <td width="40" height="11"><font color="#FF6600">金卡</font></td>
          <td width="80" height="11"><font color="#FF6600">门派高手</font></td>
          <td width="103" height="11"><font color="#FF6600">月积分排行</font></td>
          <td width="40" height="11"><font color="#FF6600">职位</font></td>
          <td width="40" height="11"><font color="#FF6600">弟子数</font></td>
          <td width="55" height="11"><font color="#FF6600">操 作</font></td>
        </tr>
        <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
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
            <div align="center"><a href="bx.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>" ><%=rs("a")%></a></div>
          </td>
          <td width="66" height="2"> 
            <div align="center"><a href="bx1.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>" title='谢谢来到快乐做客，这里永远欢迎您~'>等级</a></div>
          </td>
          <td width="40" height="2"> 
            <div align="center"><a href="bx2.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>" title='祝你财源滚滚~！'>金币</a></div>
          </td>          <td width="40" height="2"> 
            <div align="center"><a href="bx3.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>" title='祝你发大财~~'>金卡</a></div>
          </td><td width="60" height="2"> 
            <div align="center"><a href="bx4.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>" title='高手明天就是你，努力吧~！'>20大高手</a></div>
          </td>
          <td width="103" height="2"> 
            <div align="center"><a href="bx5.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>" title='祝大家玩得开心~！'>月积分排行</a></div>
          </td>
          <td width="40" height="2"> 
            <div align="center"><a href="bx6.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>" title='光大门派，需要大家！'>职位</a></div>
          </td>
          <td width="40" height="2"> 
            <div align="center"><%=rs("c")%></a></div>
          </td>
          <td width="55" height="2"> 
            <div align="center"><a href="#" onClick="window.open('mp1.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')">加入</a>|<a href="#" onClick="window.open('mp2.asp?my=<%=aqjh_name%>&amp;id=<%=rs("a")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')">离开</a></div>
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
<p align="center"><font color="#000000" size=2>　<font color="#FF0000"> 欢迎来到快乐江湖，在这里我祝大家愉快~开心!</font><br>
<font color="#0000FF">谁都想找一位好的老大照着,加入门派是没错的!</font></font>
</body>
</html>