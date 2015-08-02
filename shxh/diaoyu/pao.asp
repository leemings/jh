<%
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
Set Cn=Server.CreateObject("ADODB.Connection")
set rst=server.createobject("adodb.recordset")
diaoyu=Application("Ba_jxqy_connstr")
Cn.Open diaoyu
rst.open"select * from 钓鱼 where 姓名='"& username &"'",cn,1,1
if rst.bof then 
Response.Redirect "diaoyu.asp"
conn.close
end if
cn.execute "Delete * From 钓鱼 Where 姓名='"& username &"'"
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../chat/readonly/style.css" rel=stylesheet type="text/css">
<title></title>
</head>

<body oncontextmenu=self.event.returnValue=false bgcolor="#ffffff" text="#000000">
<div align="center">
  <p>　</p>
  <table border=1 align=center cellpadding="10" cellspacing="13" height="98" width="252">
    <tr>
      <td height="62" width="200"> 
        <table width="206">
          <tr> 
            <td valign="top" width="198"> 
              <div align="center"> 
                很可惜，你的鱼溜走了......<br><a href="diaoyu.asp">返回继续</a>
              </div>
        </table>
      </td>
    </tr>
  </table>
</div>
</body>
</html>