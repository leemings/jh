<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
'=========================================================
' Plusname: ��Ժ����������
' File: z_fy_shensu.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
Response.Buffer=True 
stats="��������"
call nav()
call head_var(0,0,"������Ժ","z_fy_fayuan.asp")
if not founduser then
Errmsg=Errmsg+"<br>"+"<li>��û���ڱ�������Ժ��Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
call dvbbs_error()
else
dim id,fyrs,fysql
id=request("id")
select case Request("action")
case ""
   fysql= "SELECT * FROM fy WHERE ID=" & id
  set fyrs=connfy.execute(fysql)
   if lcase(fyrs(2))<>lcase(membername) or fyrs(7)<>"N" then
    Errmsg=Errmsg+"<br>"+"<li>��Ǹ�������Ǳ�״����򱾰�����ᣡ"
    call dvbbs_error()
   else
   main()
   end if
case "save"
save()
end select
fyrs.close
set fyrs=nothing
connfy.close
set connfy=nothing
sub main()
response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>��������</b></td></tr><tr><td valign=middle class=tablebody2 height=100><CENTER>"
%>
<form action="z_fy_shensu.asp?id=<%=id%>&action=save" method=post  class=tablebody1>
  <table  cellspacing=1 cellpadding=3 align=center width="60%" class=tablebody1>
    <tr > 
      <td height="74" align="center">
        <table width="100%" cellspacing="1" cellpadding="3" class=tableborder1>
          <tr> 
            <td width="10%" rowspan="3" class=tablebody2>&nbsp;</td>
            <td colspan="2" class=tablebody1><b>Ͷ�����⣺</b> 
              <%=fyrs(1)%>     
            </td>     
            <td width="16%" align="right" rowspan="3" class=tablebody2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>     
          </tr>     
          <tr>      
            <td width="50%" height="30" class=tablebody1><b>ԭ�棺</b>      
              <%=fyrs(4)%>     
            </td>     
            <td width="50%" height="30" class=tablebody1><b>Ҫ��</b>      
              <%=fyrs(5)%>     
            </td>     
          </tr>     
          <tr>      
            <td colspan="2" class=tablebody1><b>Ͷ�����ݣ�</b>      
              <%=fyrs(3)%>     
            </td>     
          </tr>     
        </table>     
      </td>     
    </tr>     
    <tr>      
      <td align="center">      
             
          <b>�绤����</b><br>     
              
          <textarea name=shensu cols=80 rows=6><%=fyrs(9)%></textarea>     
        <br><br>     
        <input type=submit value="ȷ ��" name="submit" >     
        &nbsp;      
        <input type=reset value="�� д" name="reset" >     
	&nbsp;     
        <input type="button" value="�� ��" onClick="javascript:history.back()" name="button">     
       </td>     
    </tr>     
  </table>           
</form>     
<%      
end sub     
     
sub save()     
dim text     
text = replace(request.form("shensu"),"'","")     
fysql="update [fy] set bgtext='"&text&"',dateandtime=now() WHERE ID=" & id     
connfy.execute fysql     
Response.Redirect "z_fy_fayuan.asp"     
end sub     
response.write "</CENTER></td></tr></table>"          
end if       
call footer()       
%> 