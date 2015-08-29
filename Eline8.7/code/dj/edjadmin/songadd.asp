<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<meta http-equiv="Content-Language" content="zh-cn">
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
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber3" bgcolor="#FF9900" height="186">
      <tr>
        <td width="100%" height="186">
        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#333333" width="100%" id="AutoNumber4" height="100%">
<form method="POST" action="SongSave.asp">
          <tr>
            <td width="100%" align="center" bgcolor="#FFCB7D" colspan="2">舞曲添加</td>
          </tr>
          <tr>
            <td width="28%" align="right">选择分类：</td>
            <td width="72%">
              <select name="Sclassid" size="1">
                <option value="" <%if request("classid")="" then%> selected<%end if%>>选择栏目</option>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from Sclass"
rs.open sql,conn,1,1
do while not rs.eof
%>
                <option<%if cstr(request("Sclassid"))=cstr(rs("Sclassid")) and request("Sclassid")<>"" then%> selected<%end if%> value="<%=CStr(rs("SclassID"))%>" name=Sclassid><%=rs("Sclass")%></option>
<%
rs.movenext
loop
rs.close
%>
              </select></td>
          </tr>
          <tr>
            <td width="28%" align="right">舞曲名称：</td>
            <td width="72%">
            <input type="text" name="MusicName" size="26"></td>
          </tr>
          <tr>
            <td width="28%" align="right">制作人：</td>
            <td width="72%">
            <input type="text" name="WebSiteUrl" size="34" value="DJ 舞吧"></td>
          </tr>
          <tr>
            <td width="28%" align="right">舞曲链接地址：</td>
            <td width="72%">
            <input type="text" name="LF_Path" size="52"></td>
          </tr>
          <tr>
            <td width="27%" align="right">播放方式：<br>
            <font color="#FF0000">默认用RealPlayer播放</font></td>
            <td width="73%"><select size="1" name="MusicType">
            <option value="RealPlayer">RealPlayer</option>
            <option value="MediaPlayer">MediaPlayer</option>
            <option value="Flash">Flash</option>
            </select><font color="#FF0000">请一定要注意文件的格式，并选择对应的播放方式。</font></td>
          </tr>
          <tr>
            <td width="100%" align="center" colspan="2" bgcolor="#FFCB7D">
            <p align="center">
            <input type="hidden" value="add" name="act">
              <input type="submit" value=" 确 定 添 加" name="cmdok"></td>
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