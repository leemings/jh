<!--#include file="conn.asp"-->
<link href="CSS.CSS" rel="stylesheet" type="text/css">
<%
IF Session("aqjh_grade")<7 THEN
	response.write "<script>alert('��ʾ:�޸ĵ����Ϣ��Ҫ7�����Ϲ���!');window.close();</script>"
	response.end
END IF
%> 
<%
dim chinannid
chinannid=request("id")
set rs = Server.CreateObject("adodb.recordset")
sql="select *  from music where ID="+cstr(chinannid)+""
rs.open sql,conn,1,1
%>
<table align="center"   width="400">
<form action="savexg.asp?id=<%=rs("id")%>" method="post">
  <tr><td>��ţ�
        <input name="id" class="smallInput" id="id" value="<%=rs("id")%>"   disabled></td></tr>
  <tr> 
<td>�����: 
                
        <input name="name" class="smallInput" value="<%=rs("name")%>" >
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>���͸�: 
        <input name="toname" class="smallInput" value="<%=rs("toname")%>">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>��·��: 
        <input name="songurl" class="smallInput" value="<%=rs("songurl")%>"></td>
 </tr>
<tr> 
              
      <td valign="middle" >ף���� 
        <textarea name="zhufu" cols="40" rows="5" class="smallInput"><%=rs("zhufu")%></textarea><br></td>
            </tr>
            <tr> 
              <td> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input class="buttonface" type="submit" value="�޸�"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                <input class="buttonface" type="reset" value="ȡ��"></td>
            </tr>
          </form>
        </table>