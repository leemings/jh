<%PageName="Addfile"%>
<!--#include file="function.asp"-->
<%CheckAdmin1%>
<%
Classid=cstr(trim(request.QueryString("classid")))
SClassid=cstr(trim(request.QueryString("sclassid")))
NClassid=cstr(trim(request.QueryString("Nclassid")))
%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<script>
function submitonce(theform){
	//if IE 4+ or NS 6+
	if (document.all||document.getElementById){
	//screen thru every element in the form, and hunt down "submit" and "reset"
	for (i=0;i<theform.length;i++){
		var tempobj=theform.elements[i]
		if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
		//disable em
		tempobj.disabled=true
		}
	}
}
</script>
     <div align="center">
      <center>
      <table border="1" width="99%" cellspacing="0" cellpadding="0" class="TableLine" bordercolorlight="#CC0066">
        <tr>
          <td width="100%" height="20" bgcolor="#CC0066" align=center><font color="white"><b>�� �� ר ��</b></font></td>
        </tr>
        <tr>
          <td width="100%" height="22" bgcolor="#FFAAD5" align=center>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                   <form method="POST" action="AddfileSave2.asp" id=form2 name=form2 onsubmit="submitonce(this);">
                    <tr>
                      <td width="15%" align="right">���ͣ�</td>
                      <td width="85%">&nbsp;һ����Ŀ��
              <select name="classid" size="1" onchange="window.open('Addfile3.asp?classid='+this.options[this.selectedIndex].value,'_self')">
<%
set rs=server.createobject("adodb.recordset")
sql="select * from class"
rs.open sql,conn,1,1
Do while not rs.eof
%>
                <option<%if classid=cstr(rs("classid")) and classid<>"" then%> selected<%end if%> value="<%=Cstr(rs("classid"))%>"><%=rs("class")%></option>
<%
rs.MoveNext
Loop
rs.close
%>            </select>
              &nbsp;������Ŀ��
<%if classid<>"" then%>
              <select name="Sclassid" size="1" onchange="window.open('Addfile3.asp?classid=<%=request("classid")%>&sclassid='+this.options[this.selectedIndex].value,'_self')">
                <option value="" <%if request("Sclassid")="" then%> selected<%end if%>>ѡ����Ŀ</option>
<%
	sql="select * from Sclass where classid="&classid
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if Sclassid=cstr(rs("Sclassid")) and Sclassid<>"" then%> selected<%end if%> value="<%=CStr(rs("Sclassid"))%>" name=Sclassid><%=rs("Sclass")%></option>
<%
	rs.MoveNext
	Loop
	rs.close

else
%>
              <select name="Sclassid" size="1">
                <option value="" selected>ѡ����Ŀ</option>
<%end if%>
              </select>
              &nbsp;������Ŀ��
<%if sclassid<>"" then%>
              <select name="Nclassid" size="1" onchange="window.open('Addfile3.asp?classid=<%=classid%>&sclassid=<%=sclassid%>&nclassid='+this.options[this.selectedIndex].value,'_self')">
                <option value="" <%if Nclassid="" then%> selected<%end if%>>ѡ����Ŀ</option>
<%
	sql="select * from Nclass where Sclassid="&Sclassid
	rs.open sql,conn,1,1
	Do while not rs.eof
%>
                <option<%if cstr(request("Nclassid"))=cstr(rs("Nclassid")) and Nclassid<>"" then%> selected<%end if%> value="<%=CStr(rs("Nclassid"))%>" name=Nclassid><%=rs("Nclass")%></option>
<%
	rs.MoveNext
	Loop
	rs.close
%>
<%else%>
              <select name="Nclassid" size="1">
                <option value="" selected>ѡ����Ŀ</option>
<%end if%>
              </select></td>
                    </tr>
                  </table>
                </div>
           </td>
        </tr>
		<tr>
          <td width="100%" height="25" align=left>
                <div align="center">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                      <td width="15%" height="25" align="right">ר�����ƣ�</td>
                      <td width="85%"><input type="text" name="name" size="20"></td>
                    </tr>
                    <tr>
                      <td width="15%" height="25" align="right">�������ԣ�</td>
                      <td width="85%"><select name="Yuyan" size="1">
                <option value="����" selected>����</option>
                <option value="����">����</option>
                <option value="Ӣ��">Ӣ��</option>
                <option value="�պ�">�պ�</option>
                <option value="����">����</option>
                <option value="����">����</option>
                <option value="��/����">��/����</option>
                <option value="��/Ӣ��">��/Ӣ��</option>
                <option value="��/����">��/����</option>
                <option value="��/����">��/����</option>
                </select></td>
                    </tr>
                    <tr>
                      <td width="15%" height="25" align="right">��Ƭ��˾��</td>
                      <td width="85%"><input type="text" name="Gongsi" size="20"></td>
                    </tr>
                    <tr>
                      <td width="15%" height="25" align="right">�������ڣ�</td>
                      <td width="85%"><input type="text" name="Times" size="20"></td>
                    </tr>
                    <tr>
                      <td width="15%" height="25" align="right">ר��ͼƬ��</td>
                      <td width="85%"><input type="text" name="pic" size="20"></td>
                    </tr>
                    <tr>
                      <td width="15%" height="25" align="right">ר����飺</td>
                      <td width="85%"><TEXTAREA name="intro" rows=5 cols="60"></TEXTAREA></td>
                    </tr>
                    <tr>
                      <td width="15%" align="right"></td>
                      <td width="85%"></td>
                    </tr>
                    <tr>
                      <td width="15%" align="right"></td>
                      <td width="85%"></td>
                    </tr>

                  </table>
                </div>
		  </td>
        </tr>
        <tr>
          <td width="100%" height="22" bgcolor="#FFAAD5" align=center>
              <input type="hidden" value="add" name="act"> 
<%if Classid="" then%>
              <input type="button" value=" ȷ �� " name="cmdok" onclick="window.alert('����ѡ����Ŀ��')">&nbsp; 
<%elseif NClassid="" then%>
              <input type="button" value=" ȷ �� " name="cmdok" onclick="window.alert('����ѡ������Ŀ��')">&nbsp; 
<%else%>
              <input type="submit" value=" ȷ �� " name="cmdok">&nbsp; 
<%end if%>
              <input type="reset" value=" ȡ �� " name="cmdcancel"></td>
        </tr>
       </form>
      </table>
     </center>
    </div>
<%
set rs=nothing
conn.close
set conn=nothing
%></body></html>