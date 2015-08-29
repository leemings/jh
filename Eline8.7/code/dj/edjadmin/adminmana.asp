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
    <td valign=top width="600">
    <table border="0" cellpadding="3" cellspacing="4" style="border-collapse: collapse" width="100%" id="AutoNumber3" bgcolor="#FF9933">
      <tr>
        <td width="100%" bgcolor="#CC6600" align="center">
        站内管理员管理</td>
      </tr>
      <tr>
        <td width="100%">
      　<div align="center">
        <center>
      <table border="1" width="568" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#CC6600" style="border-collapse: collapse">
        <tr>
          <td width="32" align="center" bgcolor="#FFD6AC">ID号</td>
          <td width="84" align="center" bgcolor="#FFD6AC">名称</td>
          <td width="90" align="center" bgcolor="#FFD6AC">密码</td>
          <td width="98" align="center" bgcolor="#FFD6AC">最后登陆IP</td>
          <td width="144" align="center" bgcolor="#FFD6AC">最后登陆时间</td>
          <td width="36" align="center" bgcolor="#FFD6AC">停留</td>
          <td width="32" align="center" bgcolor="#FFD6AC">修改</td>
          <td width="39" align="center" bgcolor="#FFD6AC">删除</td>
        </tr>
<%
set rs=server.CreateObject("ADODB.RecordSet")
	  
sql="select * from admin"
rs.open sql,conn,1,3
%> 
<%
if rs.EOF then
%>
        <tr><td colspan=8 align=center width="562">没有用户!!!</td>
          </tr>
<%
else
	do while NOT rs.EOF
%> 
        <tr> 
          <td width="32" align="center"><%=rs("id")%>　</td>
          <td width="84" align="center"><%=rs("Username")%>　</td>
          <td width="90" align="center"><%=rs("password")%>　</td>
          <td width="98" align="center"><%=rs("LoginIP")%>　</td>
          <td width="144" align="center"><%=rs("LoginTime")%>　</td>
          <td width="36" align="center"><%=rs("LoginTimes")%>分钟</td>
          <td width="32" align="center">
          <a href="AdminModify.asp?id=<%=rs("id")%>">修改</a></td>
          <td width="39" align="center">
          <a href="Adminsave.asp?act=del&id=<%=rs("id")%>">删除</a></td>
        </tr>
<%
	rs.MoveNext
	loop
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%> 
      </table>
        </center>
      </div>
      <FORM METHOD=POST ACTION="AdminSave.asp" id=form1 name=form1>
        <p>　</p>
        <div align="center">
          <center>
        <table border="1" width="40%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#CC6600" style="border-collapse: collapse" >
          <tr> 
            <td align="center" height=20 colspan=2 bgcolor="#CC6600">添 加 管 理 员</td>
          </tr>
          <tr>
            <td align="right">管 理 员 名：</td>
            <td><input type=text name=UserName size="15"  value="" onfocus=this.select() onmouseover=this.focus() name=keyword size=14 maxlength="30"></td>
          </tr>
          <tr> 
            <td align="right">管 理 密 码：</td>
            <td><input type=text name=Password size="15"  value="" onfocus=this.select() onmouseover=this.focus() name=keyword size=14 maxlength="30"></td>
          </tr>
          <tr> 
            <td align="center" colspan=2 bgcolor="#CC6600"> 
              <input type=hidden value="add" name="act">
              <input type=submit value=增加 name="submit">
              <input type=reset name="Submit" value="取消">
            </td>
          </tr>
        </table>
          </center>
        </div>
      </FORM>

</td>
      </tr>
      <tr>
        <td width="100%" bgcolor="#CC6600">　</td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
  </center>
</div>