<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->

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
    <td valign=top width="600" bgcolor="#FF9900">
    <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" width="100%" id="AutoNumber3">

        <form method="POST" action="FriendSiteSave.asp?id=<%=id%>">
          <tr>
            <td width="100%" height="20" colspan=2 align=center bgcolor="#CC6600"><b>添 加 友 情 站 点</b></td>
          </tr>
          <tr>
            <td width="15%" align="right">站名：</td>
            <td width="85%"><input type="text" name="SiteName" size="20"></td>
          </tr>
          <tr>
            <td align="right">网址：</td>
            <td><input type="text" name="SiteUrl" size="20"></td>
          </tr>
          <tr>
            <td align="right">Logo地址：</td>
            <td><input type="text" name="LogoUrl" size="20"></td>
          </tr>
          <tr>
            <td align="right">站长：</td>
            <td><input type="text" name="SiteAdmin" size="20"></td>
          </tr>
          <tr>
            <td align="right">简介：</td>
            <td><TEXTAREA name="SiteIntro" rows=5 cols="71"></TEXTAREA></td>
          </tr>
          <tr>
            <td colspan=2 align=center bgcolor="#CC6600">
              <input type="hidden" value="add" name="act">
              <input type="submit" value=" 添 加 " name="cmdok">&nbsp; 
              <input type="reset" value=" 取 消 "  name="cmdcancel">
            </td>
          </tr>
        </form>

    </table>
    </td>
  </tr>
  </table>
  </center>
</div>