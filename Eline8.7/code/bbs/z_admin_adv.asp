<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
 if not master or session("flag")="" then
  Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
  call dvbbs_error()
 else
  dim body
  dim readme,Tlink
  call main()
  set rs=nothing
  conn.close
  set conn=nothing
 end if

 sub main()
%>
<TABLE > 
  <tr>
    <td height="30"> 
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr >
        <td><font color=blue>��ӭ<b><%=membername%></b>�������ҳ��</font>
        </td>
        </tr>
            <tr >              
          <td width="100%" valign=top>
<%
dim objFSO
dim fdata
dim objCountFile
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("save")="" then
  Set objCountFile = objFSO.OpenTextFile(Server.MapPath("text.js"),1,True)
  If Not objCountFile.AtEndOfStream Then fdata = objCountFile.ReadAll
 else
  fdata=request("fdata")
  Set objCountFile=objFSO.CreateTextFile(Server.MapPath("text.js"),True)
  objCountFile.Write fdata
 end if
 objCountFile.Close
 Set objCountFile=Nothing
 Set objFSO = Nothing
%> 
<form method=post>
            <table width="100%" border="0" cellspacing="3" cellpadding="0">
              <tr> 
                <td width="3%" height="23"> </td>
                <td width="97%" height="23">��̳������༭��<br>
                      ע�⣺�ļ������������㰲װĿ¼�µ�<font color=red>test.js</font>�ļ��������Ŀռ䲻֧��<font color=red>FSO</font>����ֱ�ӱ༭���ļ���<br>
                �������������html��js���벻�˽⣬��ֻ�޸Ĺ�����ݺͳ����ӣ�������Ҫ����༭��</td>
              </tr>
              <tr> 
                <td width="3%">  </td>
                <td width="97%"> 
                  <textarea name="fdata" cols="110" rows="20"><%=fdata%></textarea>
                </td>
              </tr>
              <tr> 
                <td width="3%"> </td>
                <td width="97%">
                  <input type="reset" name="Reset" value="�����޸�">
                  <input type="submit" name="save" value="�ύ�޸�"> ��ǰ�ļ�·����<b><%=Server.MapPath("text.js")%></b>
                </td>
              </tr>
            </table></form>
            <p> </p>
            </td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%end sub%>