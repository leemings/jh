<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<%
id=request.QueryString("id")
set rs=server.createobject("adodb.recordset")
sql="select * from FriendSite where id="&id
rs.open sql,conn,1,1
if rs.eof then
	errmsg=errmsg+"<br>"+"<li>操作错误！请联系管理员"
	call error()
	Response.end
else
	SiteName=rs("SiteName")
	SiteUrl=rs("SiteUrl")
	SiteIntro=rs("SiteIntro")
	LogoUrl=rs("LogoUrl")
	SiteAdmin=rs("SiteAdmin")
end if
rs.close

%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" id="AutoNumber2">
    <tr>
      <td width="100%"><img border="0" src="补间.gif" width="1" height="3"></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
<table border="0" width="760" cellspacing="0" cellpadding="0" style="border-collapse: collapse">
  <tr>
    <td valign=top width=150>
    <!--#include file="admin_left.asp"-->
    　</td>
    <td valign=top width=10>　</td>
    <td valign=top width="600">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber3" bgcolor="#FF9900">
      <tr>
        <td width="100%">
      <table border="0" width="100%" cellspacing="3" cellpadding="3" class="TableLine" style="border-collapse: collapse">
        <form method="POST" action="FriendSiteSave.asp?id=<%=id%>">
          <tr>
            <td width="100%" height="20" colspan=2 align=center bgcolor="#CC6600"><b>修 改 友 情 站 点</b></td>
          </tr>
          <tr>
            <td width="15%" align="right">站名：</td>
            <td width="85%"><input type="text" name="SiteName" value="<%=SiteName%>" size="20"></td>
          </tr>
          <tr>
            <td align="right">网址：</td>
            <td><input type="text" name="SiteUrl" value="<%=SiteUrl%>" size="20"></td>
          </tr>
          <tr>
            <td align="right">Logo地址：</td>
            <td><input type="text" name="LogoUrl" value="<%=LogoUrl%>" size="20"></td>
          </tr>
          <tr>
            <td align="right">站长：</td>
            <td><input type="text" name="SiteAdmin" value="<%=SiteAdmin%>" size="20"></td>
          </tr>
          <tr>
            <td align="right">简介：</td>
            <td><TEXTAREA name="SiteIntro" rows=5 cols="68"><%=SiteIntro%></TEXTAREA></td>
          </tr>
          <tr>
            <td colspan=2 align=center bgcolor="#CC6600">
              <input type="hidden" value="edit" name="act">
              <input type="submit" value=" 修 改 " name="cmdok">&nbsp; 
              <input type="reset" value=" 清 除 "  name="cmdcancel">
            </td>
          </tr>
        </form>
      </table>
</td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
  </center>
</div>

<%
set rs=nothing
conn.close
set conn=nothing%>