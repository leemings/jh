<!--#include file="function.asp"-->
<%CheckAdmin1%>
<!--#include file="conn.asp"-->
<%
Specialid=request.QueryString("Specialid")
set rs=conn.execute("delete from Special where Specialid="&Specialid)
conn.close
set conn=nothing
classid=request.QueryString("classid")
Nclassid=request.QueryString("Nclassid")
page=request.QueryString("page")
 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="RSHOP" content="k666影音世界http://www.vv66.com">
<meta Author="Recall Star" content="k666影音世界http://www.vv66.com">
<title>k666音乐屋</title>
<!--#include file="style.asp"-->
</head>
<body topmargin="111" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellspacing="0" width="60%">
    <tr>
      <td width="100%" bgcolor="#CC0066">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td width="100%" bgcolor="#FFFFFF" height="80" align="center">
                <b>O K !&nbsp; 操 作 完 成 !&nbsp; ^_^</b>
                <p><b><a href="javascript:history.go(-1)">...::: 点 此 返 回 
                :::...</a></b>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
</body>                    
</html>           


