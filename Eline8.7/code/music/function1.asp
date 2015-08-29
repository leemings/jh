<%
sub error()
%>

<div align="center">
  <center>
  <table border="0" cellspacing="0" width="60%">
    <tr>
      <td width="100%" bgcolor="#56B0F4">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
                <td width="100%" bgcolor="#FFFFFF" height="80" align="center"> 
                  <p>只有本站会员才能视听歌曲！ <%=errmsg%>！&nbsp; :( </p>
                  <p><A 
                  href="User/Reg.asp" target="_blank">点此注册会员</A></p>
                  <p><a href="index.asp" target="_blank">回到首页登陆</a></p>
                </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
                
       
<%
end sub
%>