<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：想给朋友点歌请先进入江湖聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<%
id=request("id")
if id="" then
	errmsg="你尚未选择歌曲!"
	call error()
	Response.End 
end if
%>
<%
set rs=conn.execute("SELECT * FROM MusicList where id="+id)
if rs.eof then
	errmsg="歌曲选择错误。请重选!"
	call error()
	Response.End 
else
	MusicName=rs("MusicName")
	Singer=rs("Singer")
	Wma=rs("Wma")
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>点歌：<%=Singer%>--{<%=MusicName%>}, </title>
<link rel=stylesheet href=images/style.css>
</head>
<body bgcolor="#B4DEF8" topmargin="0" leftmargin="0">             
<table border="1" width="100%" cellspacing="0" bordercolor="#56B0F4" bordercolordark="#56B0F4"> 
  <tr> 
    <td width="100%" height=22 Bgcolor="#56B0F4" align="center"><b>EMAIL发送点歌单</b></td>
  </tr>
  <tr>
    <td width="100%">               
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <form name="form1" method="post" action="Sendmail.asp">   
          <tr>
            <td width=50% bgcolor="#B4DEF8">&nbsp;&nbsp; <b>送给谁</b>：<INPUT type=text name=name size=20 value='您朋友的名字' onFocus=this.value=''></td>
            <td width=50% bgcolor="#B4DEF8">&nbsp;<b>Email</b>： <INPUT type=text name=mail size=20 value='您朋友的信箱' onFocus=this.value=''></td>
          </tr>
          <tr>
            <td width=50% bgcolor="#B4DEF8">&nbsp;&nbsp; <b>点歌人</b>：<INPUT type=text name=cname size=20 value='您的名字' onFocus=this.value=''></td>
            <td width=50% bgcolor="#B4DEF8">&nbsp;<b>Email</b>： <INPUT type=text name=cmail size=20 value='您的信箱' onFocus=this.value=''></td>
          </tr>
          <tr>
            <td width=100% valign=top colspan="2" bgcolor="#B4DEF8" align="center"><TEXTAREA name=dmail rows=4 cols="60">祝福...</TEXTAREA></td>
          </tr>
          <tr>
            <td width=100% height=23 bgcolor="#56B0F4" align=center colspan="2"><input type="submit"  value="提交" name="Submit"> 
              &nbsp;&nbsp; <input type="reset"  value="重写" name="B1"></td>
          </tr>			
      </table><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<input type="text" name="email" value="<%=Singer%>--{<%=MusicName%>}" class=kuang size="42" maxlength="50">
      <input type="text" name="fmail" value="http://zhzx.jjedu.org/eline/music/MusicPlay.asp?id=<%=id%>"　class=kuang size="42" maxlength="50">
    </td>
  </FORM></tr>
</table>
</body>
</html>