<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from �û� where ����='" & sjjh_name &"'",conn,3,3
mygj=rs("�ȼ�")*sjjh_gjsx+4500 
myfy=rs("�ȼ�")*sjjh_fysx+3000
if rs("����")>mygj then
	conn.execute("update �û� set ����="& mygj &" where ����='"&sjjh_name&"'")
end if
if rs("����")>myfy then
	conn.execute("update �û� set ����="& myfy &" where  ����='"&sjjh_name&"'")
end if
if rs("��Ա")=true and clng(DateDiff("d",now(),rs("��Ա����")))>0 then
	pdstr="<font size=-1>[�ݵ��ƻ�Ա]"&rs("��Ա����")&"</font>"
else
	pdstr=""
end if
wgj=rs("�书��")
nlj=rs("������")
tlj=rs("������")
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<title>���ֽ��� - [<%=sjjh_name%>]״̬</title>
<body style="BACKGROUND-COLOR: buttonface; BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; BORDER-RIGHT: medium none; BORDER-TOP: medium none; PADDING-BOTTOM: 6px; PADDING-LEFT: 6px; PADDING-RIGHT: 6px; PADDING-TOP: 6px" marginwidth="0" marginheight="0" leftmargin="0" topmargin="10" LANGUAGE="javascript"  >
<table width="280" border="1" cellspacing="0" cellpadding="0" class="p150" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="left">
  <tr valign="top" align="left"> 
    <td height="250" colspan="4" class="p9"> 
      <!--#include file="../z_showvisualmy.asp"-->
      <p>&nbsp;</p></td>
  </tr>
  <tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font color="#000000" size="2">����:</font></td>
    <td width="26%"><font size="2"><%=rs("����")%></font>��</td>
    <td width="13%"><font color="#000000" size="2">ְҵ:</font></td>
    <td width="26%"><font size="2"><%=rs("ְҵ")%></font>��</td>
  </tr>
  <tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font color="#000000" size="2">ʦ��:</font></td>
    <td width="26%"><font size="2"><%=rs("ʦ��")%></font>��</td>
    <td width="13%"><font color="#000000" size="2">����:</font></td>
    <td width="26%"><font size="2"><%=rs("����")%></font>��</td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9" height="8"><font color="#000000" size="2">����:</font></td>
    <td width="26%" height="8"><font size="2"><%=rs("����")%></font></td>
    <td width="13%" height="8"><font color="#000000" size="2">����:</font></td>
    <td width="26%" height="8"><font size="2"><%=rs("����")%></font></td>
  </tr>
  <%if rs("��Ա�ȼ�")>0 then%>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font color="#000000" size="2">��Ա:</font></td>
    <td width="26%"><font size="2"><%=rs("��Ա�ȼ�")%></font>��</td>
    <td width="13%"><font color="#000000" size="2">����</font></td>
    <td width="26%"><font size="2"><%=rs("��Ա����")%></font>��</td>
  </tr>
  <%end if%>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font color="#000000" size="2">ɱ��:</font></td>
    <td width="26%"><font size="2"><%=rs("ɱ����")%><font color="#FF0000">/</font><font size="2" color="#FF0000"><b><%=int(Application("sjjh_killman"))%></b></font></font></td>
    <td width="13%"><font color="#000000" size="2">����:</font></td>
    <td width="26%"><font size="2"><%=rs("��ɱ��")%>��</font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font size="2" color="#000000">�ȼ�: </font></td>
    <td width="26%"><font size="2"><%=rs("�ȼ�")%></font>��</td>
    <td width="13%"><font size="2" color="#000000">����: </font></td>
    <td width="26%"><font size="2" color="#000000"><%=rs("grade")%></font>��</td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9" height="2"><font size="2" color="#000000">����:</font></td>
    <td width="26%" height="2" nowrap><font size="2"><%=rs("����")%></font></td>
    <td width="13%" height="2" nowrap><font size="2" color="#000000">���: </font></td>
    <td width="26%" height="2" nowrap><font size="2"><%=rs("���")%></font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font size="2" color="#000000">���: </font></td>
    <td width="26%"><font size="2"><%=rs("allvalue")%></font>��</td>
    <td width="13%"><font size="2" color="#000000">����:</font></td>
    <td width="26%"><font size="2"> 
      <%if rs("grade")=3 and rs("���")="����" then%>
      <%=rs("�ݶ�����")-int(rs("�ݶ�����")/500)*500%><font color="#FF0000">/500</font> 
      <%else%>
      <%=rs("�ݶ�����")-int(rs("�ݶ�����")/1000)*1000%><font color="#FF0000">/1000</font> 
      <%end if%>
      </font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font size="2" color="#000000">����: </font></td>
    <td width="26%"><font size="2"><font color=red> 
      <%if rs("grade")=3 and rs("���")="����" then%>
      <%=int(rs("�ݶ�����")/500)%> 
      <%else%>
      <%=int(rs("�ݶ�����")/1000)%> 
      <%end if%>
      </font></font>��</td>
    <td width="13%"><font size="2" color="#000000">����: </font></td>
    <td width="26%"><font size="2"> 
      <%if rs("����")=true then%>
      �������� 
      <%else%>
      û�б��� 
      <%end if%>
      </font></td>
  </tr>
  <%
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<=60 then%>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td colspan="4" class="p9"><div align="center"> <font size="2" color="#000000">������ʣ:</font><font size="2" color=red><b><%=60-sj%>��</b></font></div></td>
  </tr>
  <%end if%>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%"><font size="2" color="#000000">��:</font></td>
    <td width="26%" height="20"><font color="#0000FF"><font color=red size="2"><%=rs("��Ա��")%>Ԫ 
      </font></font></td>
    <td width="13%" height="20"><font size="2" color="#000000">���:</font></td>
    <td width="26%" height="20"><font color="#0000FF"><font color=red size="2"><%=rs("���")%>��</font></font></td>
  </tr>
  <tr  bgcolor=""onmouseover="this.bgColor='#DFEFFF';" onmouseout="this.bgColor='';"> 
    <td><font size="2" color="#000000">����:</font></td>
    <td height="24"><font color="#0000FF"><font size="2"> 
      <%
if rs("�ȼ�")<5  then response.write("����է��")
if rs("�ȼ�")>=5 and rs("�ȼ�")<10  then response.write("��������")
if rs("�ȼ�")>=10 and rs("�ȼ�")<15  then response.write("С�гɾ�")
if rs("�ȼ�")>=15 and rs("�ȼ�")<20  then response.write("�����Ժ�")
if rs("�ȼ�")>=20 and rs("�ȼ�")<35  then response.write("�д�����")
if rs("�ȼ�")>=35 and rs("�ȼ�")<45  then response.write("һ������")
if rs("�ȼ�")>=45 and rs("�ȼ�")<55  then response.write("��������")
if rs("�ȼ�")>=55 and rs("�ȼ�")<65  then response.write("��������")
if rs("�ȼ�")>=65 and rs("�ȼ�")<75  then response.write("������ʤ")
if rs("�ȼ�")>=75 and rs("�ȼ�")<80  then response.write("�����ɵ�")
if rs("�ȼ�")>=80 and rs("�ȼ�")<85  then response.write("С��")
if rs("�ȼ�")>=85 and rs("�ȼ�")<90  then response.write("����")
if rs("�ȼ�")>=90 and rs("�ȼ�")<95  then response.write("���۴���")
if rs("�ȼ�")>=95 and rs("�ȼ�")<100  then response.write("����")
if rs("�ȼ�")>=100 then response.write("����")%>
      </font></font></td>
    <td height="24"><font color="#000000" size="2">���:</font></td>
    <td height="24"><font color="#0000FF"><font size="2"><%=rs("���")%></font><font color=red size="2"> 
      </font></font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9" height="12"><font size="2" color="#000000">����: 
      </font></td>
    <td width="26%" height="12"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("����")%></font></font> 
    </td>
    <td width="13%" height="12"><font size="2" color="#000000">����: </font></td>
    <td width="26%" height="12"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("����")%></font> 
      </font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%"><font size="2" color="#000000">����: </font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("����")%></font> 
      </font></td>
    <td width="13%" height="24"><font size="2" color="#000000">֪��: </font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("֪��")%></font> 
      </font></td>
  </tr>
  <tr  bgcolor="" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td width="13%"><font size="2" color="#000000">����: </font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("����")%></font> 
      </font></td>
    <td width="13%" height="24"><font size="2" color="#000000">����:</font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("sl")%></font></font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%"><font size="2" color="#000000">��ż: </font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("��ż")%></font> 
      </font></td>
    <td width="13%" height="24"><font size="2" color="#000000">����:</font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("����")%></font></font></td>
  </tr>
  <%if rs("��ż")<>"��" then%>
  <tr  bgcolor="" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td colspan="4" class="p9"> 
      <div align="center"><font size="2" color="#000000">����:</font><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("������")%>��</font></font><font color="#0000FF" size="2"> 
        </font><font size="2" color="#000000">����:<%=rs("��������")%> ���ʱ��:<%=DateDiff("d",rs("��������"),date())%>��</font></div>
    </td>
  </tr>
  <%end if%>
  <tr> 
    <td colspan="4"> <div align="center"><font size="2" color=red><b><%if sjjh_sx=1 then%>���޴�<%else%>���޹ر�<%end if%></b></font><br><font size="2" color="#000000">����<br>
        </font><img height=12
<%if int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*200)<200 then
abc=int(int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*200)/50)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*200)%>><img height=12 src="img/5.gif" width=<%=200-int((rs("����")/(rs("�ȼ�")*sjjh_tlsx+5260+tlj))*200)%>> 
        <%else%>
        src="img/4.gif" width=200> 
        <%end if%>
        <br>
        <font size="2"><%=rs("����")%><font color="#FF0000"><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_tlsx+5260+tlj%><%end if%></font></font></div></td>
  </tr>
  <tr> 
    <td colspan="4" class="p9" height="23"> <div align="center"><font size="2" color="#000000">����<br>
        </font><img height=12
<%if int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+nlj))*200)<200 then
abc=int(int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+nlj))*200)/50)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+nlj))*200)%>><img height=12 src="img/5.gif" width=<%=200-int((rs("����")/(rs("�ȼ�")*sjjh_nlsx+2000+nlj))*200)%>> 
        <%else%>
        src="img/4.gif" width=200> 
        <%end if%>
        <br>
        <font size="2"><%=rs("����")%><font color="#FF0000"><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_nlsx+2000+nlj%><%end if%></font></font></div></td>
  </tr>
  <tr> 
    <td colspan="4" class="p9"> <div align="center"><font size="2" color="#000000">�书<br>
        </font><img height=12
<%if int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+wgj))*200)<200 then
abc=int(int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+wgj))*200)/50)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+wgj))*200)%>><img height=12 src="img/5.gif" width=<%=200-int((rs("�书")/(rs("�ȼ�")*sjjh_wgsx+3800+wgj))*200)%>> 
        <%else%>
        src="img/4.gif" width=200> 
        <%end if%>
        <br>
        <font size="2"><%=rs("�书")%><font color="#FF0000"><%if sjjh_sx=1 then%>/<%=rs("�ȼ�")*sjjh_wgsx+3800+wgj%><%end if%></font></font></div></td>
  </tr>
</table>
<%rs.close
set rs=nothing
conn.close
%>