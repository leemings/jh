<%
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sqlstr="SELECT ��ż FROM �û� where ����='"&username&"'"
Set Rs=conn.Execute(sqlstr)
peiou=rs("��ż")
rs.Close        
set rs=nothing  
if peiou="��" then
Response.Write "<script language=javascript>alert('�㻹û����ż��������������');history.back();</script>"
else
%>
<HTML><HEAD><TITLE>�¶�Ժ</TITLE><LINK 
href="../../style.css" rel=stylesheet type=text/css>
</HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=0 oncontextmenu=self.event.returnValue=false>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=548 height="380">
  <TBODY> 
  <TR>
    <TD background=jh/history_table_bg.gif height=377 vAlign=top width="556"> 
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain 
      width=531>
        <TBODY> 
        <TR>
          <TD vAlign=top bgcolor="#000000" width="552">
            <TABLE align=center border=0 
            cellPadding=0 cellSpacing=0 class=p9 width="90%">
              <TBODY> 
              <TR>
                <TD align=middle vAlign=center> 
                  <P class=p9><font color="#FF0000"><b><img border="0" src="../../image/052.gif">��������ʾ��<img border="0" src="../../image/052.gif"></b></font><br>
                  ����ǽ��������Ĺ¶�Ժ�����ڴ��ڽ������������Թ¶�Ժ��Ĺ¶��ܶ࣬������������10��<br>
                    �����ӾͿ�����һ���¶��ؼҡ���ÿ��ι�������չ�������ĺ��ӳ���ᱨ����ġ��ȵ�һ��ʱ�������� ������׬��ҵġ�                                       
                  </P>                                   
                  <table width="550" border="0" cellspacing="0" cellpadding="0" align="center">                                   
                    <tr>                                    
          <TD vAlign=top align=middle width=16 rowSpan=8><IMG                                  
            src="../../image/dragon.gif"><BR><IMG height=152                                  
            src="../../image/line.gif" width=1><BR><BR><IMG height=35                                  
            src="../../image/arrow.gif" width=10> </TD>                                 
                      <td valign="top" width="476">                                    
                        <div align="center" style="width: 507; height: 216">                                    
                          <div align="center">                                    
                            <div align="center">                                    
                              <table border="0" width="468" cellspacing="1" cellpadding="0"                                   
            height="20">                                   
                                <center>                                   
                                </center>                                   
                              </table>                                   
                            </div>                                   
                          </div>                                   
                          <div align="center">                                    
                            <center>                                   
                              <table border="0" width="509" cellspacing="1" cellpadding="0" class="p9">                                   
                                <tr>                                    
                                  <td width="503">                                    
                                    <table border="0" width="105%" cellspacing="1" cellpadding="0">                                   
                                      <tr>                                    
                                        <td class="p2" width="100%">                                    
                                          <p align="center"><font color="#FFCC00" class="p9">�����¶�Ժ��Ĺ¶�:</font><img border="0" src="../../image/251.GIF" width="23" height="20"></p>                                  
                                        </td>                                  
                                      </tr>                                  
                                      <tr>                                   
                                        <td class="p9" width="100%">��-�����¶�</td>                                  
                                      </tr>                                  
                                    </table>                                  
                                    <table border="0" width="506" cellspacing="1" cellpadding="0">                                  
                                      <tr>                                   
                                        <td class="p9" width="168" height="20">��-Ϊ��ĺ���ȡһ�����������ְɣ�)</td>                                  
                                        <td class="p3" width="335" rowspan="2">                                   
                                          <form name="form1" method="POST" action="buysheep.asp">                                  
                                            <input type="text" name="sheepname" size="33"                                  
                        style="font-family: ����; font-size: 12px">                                  
                                            <input                                  
                        type="button" value="����" name="B1"                                  
                        style="font-family: ����; font-size: 12px">                                  
                                          </form>                                  
                                          <script language="vbscript">                                   
<!--                                   
sub b1_onclick                                   
if form1.sheepname.value="" then                                   
msgbox"С�������ֲ���Ϊ��"                                   
else                                   
form1.submit                                   
end if                                   
end sub                                   
-->                                   
</script>                                  
                                        </td>                                  
                                      </tr>                                  
                                      <tr>                                   
                                        <td class="p3" width="168" height="20"></td>                                  
                                      </tr>                                  
                                    </table>                                  
                                  </td>                                  
                                </tr>                                  
                              </table>                                  
                            </center>                                  
                          </div>                                  
                          <!--#include file="data1.asp"-->                                  
                          <p><font color="#FFCC00"><span class="p9">�㹤���ˣ�</span>                                   
                            <%rs.open("select * from sheep where user='"&username&"'"),conn,1,1                                                   
if  rs.bof then%>                                    
                            <span class="p9">�������С��Ϊ��׬Ǯ����</span>                                         
                            <%else                                                   
 if rs("feedsheepday")>=3then%>                                    
                            </font></p>                                    
                          <form method="POST" action="sellmilk.asp">                                    
                            <p><span class="p9">�㺢��Ϊ���ڹ¶�Ժ�����<%=rs("milk")%>������λ��������¶�Ժ�ն���Сʱ�Ĺ�Ǯ�أ�                                     
                              <input                                       
          type="text" name="milk" size="9"                                       
          style="font-family: ����; font-size: 12px">                                    
                              <input type="submit"                                       
          value="�չ�Ǯ��" name="B2"                                       
          style="font-family: ����; font-size: 12px">                                    
                              <input type="hidden"                                       
          name="sheepname" value="<%=rs("sheepname")%>">                                    
                              <%else%>                                    
                              ��ĺ��ӻ�û�ҵ�������</span>��</p>                                    
                          </form>                                    
                          <%end if%>                                    
                                                              
                            <%end if%>                                    
                        </div>                                    
                      </td>                                    
          <TD vAlign=top align=middle width=16 rowSpan=8><IMG                                   
            src="../../image/dragon.gif"><BR><IMG height=152                                   
            src="../../image/line.gif" width=1><BR><BR><IMG height=35                                   
            src="../../image/arrow.gif" width=10> </TD>                                  
                    </tr>                                    
                  </table>                                    
                </TD></TR></TBODY></TABLE>                                    
          </TD></TR></TBODY></TABLE></TD></TR>                                    
  </TBODY></TABLE></BODY></HTML>                    
  <%                   
end if                                    
set rs=nothing                   
conn.Close                   
set conn=nothing                   
%>                   
                                  
