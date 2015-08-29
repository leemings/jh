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
                <td width="100%" bgcolor="#FFFFFF" height="80" align="center"> Error£¡&nbsp; 
                  <%=errmsg%>£¡&nbsp; :( 
                  <p><a href="javascript:history.go(-1)">...::: µã ´Ë ·µ »Ø :::...</a> 
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