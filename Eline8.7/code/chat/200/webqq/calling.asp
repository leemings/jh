<%
  '��ȡ������Ϣ��ʱ��
  strHour = Hour(Time())
  If Len(strHour) = 1 then strHour = "0" & strHour
  strMinute = Minute(Time())
  If Len(strMinute) = 1 then strMinute="0" & strMinute
  strTime = "" & strHour & "ʱ" & strMinute & "��"
  
  '��ȡ��Ϣ����
  MyMsg = Request("InputMsg")
  If MyMsg <> "" Then
    Receiver = Request("Receiver")
    UserMsg = Receiver & "Msg"
    Application.Lock
'    Application(UserMsg) = "[" & Receiver & "]" & Session("rName") & "��" & strTime & "������Ϣ������<BR>" & MyMsg
    Application(UserMsg) = "  <a href=onemam.asp?bname=" & Session("rName") & "  target=_blank>" & Session("rName") & "</a>��" & strTime & "������Ϣ������<BR>&nbsp;&nbsp;      " & MyMsg

' target="_blank"
    Application.Unlock
'    Response.Redirect "Sendmsg.asp"
  END IF
%>
<HTML><HEAD><TITLE>���ߴ�Ѷ</TITLE>
<style type="text/css">
<!--
.link1 {  color: #FFFFFF}
.link2 {  color: #FF0000}
.link3 {  color: #006600}
.big1 {  font-size: 14px; text-align: justify; vertical-align: 1%; line-height: 18px}
.title {  font-size: 14px; background-color: #CCFFCC; font-weight: bold}
.botton {  color: #000066; background-color: #FFCC99; font-size: 12px; cursor: hand;}
-->
</style>
<link rel="stylesheet" href="../3508.css">
</HEAD>
<BODY  bgcolor="#ffffdf">
<FORM ACTION="Calling.asp">
  <div align="center"> 
    <table width="100%" border="0" cellspacing="0">
      <tr>
        <td>
          <div align="center">�������룺</div>
        </td>
      </tr>
    </table>
    <table width="100%" border="0">
      <tr> 
        <td colspan="2"> 
          <div align="right"></div>
          <div align="center">
            <select name="Receiver" class="botton">

              <% For I = 0 To (Application("TotalUsers") - 1) %> 
              <option value="<%= Application("OnlineUser")(I) %>"> <%= Application("OnLineUser")(I) %></option>
              <% Next %> 
            </select>
          </div>
        </td>
      </tr>
      <tr> 
        <td colspan="2"> 
          <div align="center"> 
            <input  class="buttonface" type="TEXT" name="InputMsg" size="20">
          </div>
        </td>
      </tr>
    </table>
    <input  class="botton"  type="SUBMIT" value="����" name="SUBMIT">
    <input  class="botton" type="RESET" value="���" name="RESET">
   
  </div>
  </FORM>          
  </BODY>          
</HTML>          
       
       
     
     
