<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update fw set yhid=0 where zysj<now()-2"
if DateDiff("d",xhday,now())<=10 and mysex="��" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ӳ���ǰ10��ֻ����ĸ�׸�����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if

rs.open "select * from fw",conn
%>
<HTML>
<HEAD>
<TITLE>��������С��</TITLE>
<META http-equiv=""Content-Type"" content=""text/html"">
<link rel="stylesheet" href="../../css.css">
</HEAD>

<body bgcolor="#8EB4D9" background="../../bg.gif">
<div align="center">
        <table width="531" border="1" cellspacing="0" cellpadding="0" height="56">
          <tr> 
            <td width="34" height="28"> 
              <div align="center">����</div>
            </td>
            <td width="106" height="28"> 
              <div align="center">��������</div>
            </td>
            <td width="97" height="28"> 
              <div align="center">������</div>
            </td>
            <td width="89" height="28"> 
              <div align="center">��ż</div>
            </td>
            <td width="112" height="28"> 
              <div align="center">�۸�</div>
            </td>
            <td width="79" height="28"> 
              <div align="center">����</div>
            </td>
          </tr>
          <%do while not(rs.eof or rs.bof)%>
          <tr> 
            <td width="34" height="28"> 
              <div align="center"><%=rs("id")%></div>
            </td>
            <td width="106" height="28"> 
              <div align="center"><%=rs("a")%></div>
            </td>
            <%yhid=rs("yhid")
			yhxm="��"
			poxm="��"
			if yhid<>0 then
				sql="select ����,��ż from �û� where id="&yhid
				'rs1.open "select ����,��ż from �û� where id="&yhid,conn
				set rs1=conn.execute(sql)
				if rs1.bof or rs1.eof then
					conn.execute "update fw set yhid=0,poid=0 where yhid="&yhid
					yhid=0
				else
					yhxm=rs1("����")
					poxm=rs1("��ż")
				end if
				rs1.close
				set rs1=nothing
			end if%>
            <td width="97" height="28"> 
              <div align="center"><%=yhxm%></div>
            </td>
            <td width="89" height="28"> 
              <div align="center"><%=poxm%></div>
            </td>
            <td width="112" height="28"> 
              <div align="center"><%=rs("jgfs")%>��<%=rs("b")%></div>
            </td>
            <td width="79" height="28"> 
              <div align="center">
                <%if rs("yhid")=0 then%>
	                <a href="zyfw.asp?id=<%=rs("id")%>">���÷���</a> 
                <%else
					  if poxm=aqjh_name then%>
                			<a href="love.asp?id=<%=rs("id")%>">���뷿��</a> 
		             <%else%>
        			        �Ѿ����� 
		             <%end if
				end if%>
              </div>
            </td>
          </tr>
		  <%rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		  conn.close
		  set conn=nothing%>
        </table><br>
<center>�ڿ��ֽ�����վ��</center>
      </TD>
</form>
</body>
</html>