<!--#include file="data.asp"--> 
<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
sql="select * from ���� where ����='"& username &"'"
Set Rs=connt.Execute(sql)
If rs.EOF Or rs.bof Then
   Response.Redirect "diaoyu.asp"
   connt.close
End If
sql="Delete * From ���� Where ����='"& username &"'"
connt.execute sql
connt.close
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../../style.css" rel=stylesheet type="text/css">
<title></title>
</head>
<body oncontextmenu=self.event.returnValue=false background='../../image/bg.gif'>
<div align="center">
  <p>��</p>
  <table border=1 align=center cellpadding="10" cellspacing="13" height="98" width="252">
    <tr>
      <td height="62" width="200"> 
        <table width="206">
          <tr> 
            <td valign="top" width="298"> 
              <div align="center"> 
                �ܿ�ϧ�������������......<br><a href="diao.asp">���ؼ���</a>
              </div>
        </table>
      </td>
    </tr>
  </table>
</div>
</body>
</html>