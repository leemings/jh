<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<%
sclassid=trim(request("sclassid"))
addnum=trim(request("addnum"))
MusicName=request("MusicName")
DJUser=request("DJUser")
LF_Path=request("LF_Path")
MusicType=request("MusicType")
%>
<meta http-equiv="Content-Language" content="zh-cn">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" id="AutoNumber2">
    <tr>
      <td width="100%"><img border="0" src="����.gif" width="1" height="3"></td>
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
    ��</td>
    <td valign=top width=10>��</td>
    <td valign=top width="600">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber3" bgcolor="#FF9900" height="186">
      <tr>
        <td width="100%" height="186">
        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#333333" width="101%" id="AutoNumber4" height="1">
 <form method="POST" action="songsave2.asp">
          <tr>
            <td width="100%" align="center" bgcolor="#FFCB7D" colspan="6" height="33">�����������
            �ڶ���(�������)</td>
          </tr>
          <tr>
            <td width="9%" height="1">���</td>
            <td width="5%" height="1">��������</td>
            <td width="14%" height="1">��������</td>
            <td width="17%" height="1">
              ������</td>
            <td width="33%" height="1">
              ����·��</td>
            <td width="22%" height="1">
              ���ŷ�ʽ</td>
          </tr>
<%for j=1 to addnum%>
          <tr>
            <td width="9%" height="1"><%=j%></td>
            <td width="5%" height="1">
              <select name="Sclassid<%=j%>" size="1">
                <option value="" <%if request("classid")="" then%> selected<%end if%>>
                ѡ����Ŀ</option>
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
            <td width="14%" height="1">
            <input type="text" name="MusicName<%=j%>" size="9" value="<%=MusicName%>"></td>
            <td width="17%" height="1">
              <input type="text" name="DJUser<%=j%>" size="11" value="<%=DJUser%>"></td>
            <td width="33%" height="1">
            <input type="text" name="LF_Path<%=j%>" size="20" value="<%=LF_Path%>"></td>
            <td width="22%" height="1">
              <select size="1" name="MusicType<%=j%>">
            <option value="RealPlayer" <%if musictype="RealPlayer" then%>selected<%end if%>>RealPlayer</option>
            <option value="MediaPlayer" <%if musictype="MediaPlayer" then%>selected<%end if%>>MediaPlayer</option>
            <option value="Flash" <%if musictype="Flash" then%>selected<%end if%>>Flash</option>
            </select></td>
          </tr>
  <%next%>
            <tr>
            <td width="100%" align="center" colspan="6" bgcolor="#FFCB7D" height="41">
            <p align="center">
            <input type="hidden" value=<%=addnum%> name="addnum">
              <input type="submit" value="������������������ݿ�" name="cmdok"></td>
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