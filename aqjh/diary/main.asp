<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if request("name")="" then
name=aqjh_name
else
name=request("name")
end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From �ռ��û� where ����='"& name &"'", conn1, 1,1
if rs1.eof and rs1.bof then
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
Response.Redirect "reg.asp"
end if
rjname=rs1("�ռǱ���")
lys=rs1("������")
jds=rs1("����")
xhs=rs1("�ʻ�")
sm=rs1("˵��")
wcs=rs1("�Ĳ�")
dim mylb(3)
mylb(1)=rs1("lb1")
mylb(2)=rs1("lb2")
mylb(3)=rs1("lb3")
rjs=rs1("�ռ���")
jl=rs1("��������")
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
%>
<html><head><title><%=name%>���ռǱ�</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head><LINK href="images/diary.css" rel=stylesheet type=text/css><body bgcolor="#ebe4de" text="#FFFFFF">
<DIV ID="overDiv" STYLE="position:absolute; visibility:hide; z-index: 1;"></DIV>
<SCRIPT LANGUAGE="JavaScript" src="images/overlib.js"></SCRIPT>
<table align=center border=0 cellpadding=0 cellspacing=0 width=574>
<tr><td><img height=14 src="images/top-1.gif" width=574></td></tr><tr>
    <td align=middle background=images/top1-1.gif height="2">
      <table border=0 cellpadding=2 cellspacing=0 width="100%">
        <tr> 
          <td align=middle width="1%">&nbsp;</td>
          <td align=middle width="20%"> ��<font color=blue><%=rjname%></font>��</td>
          <td width="11%" rowspan="7">&nbsp;</td>
          <td valign=top width="63%" rowspan="7"> 
            <table border=0 cellpadding=0 cellspacing=0 width="100%">
              <tr align=right> 
                <td><a href="list.asp"><img alt=�����ռ� border=0 height=21 src="images/anniu2.gif" width=80></a></td>
                <td><a href="top.asp?fs=1"><img border=0 height=21 src="images/anniu3.gif" width=91></a></td>
                <td><a href="top.asp?fs=2"><img alt=�ռ����� border=0 height=21 src="images/anniu4.gif" width=78></a></td>
              </tr>
            </table>
            <div align="center">��<a href="main.asp?fs=1&name=<%=name%>" ><font color=blue>ȫ��</font></a>����<a href="main.asp?fs=2&name=<%=name%>&zz=1" ><font color=blue><%=mylb(1)%></font></a> 
              ����<a href="main.asp?fs=2&name=<%=name%>&zz=2" ><font color=blue><%=mylb(2)%></font></a> 
              ����<a href="main.asp?fs=2&name=<%=name%>&zz=3" ><font color=blue><%=mylb(3)%></font></a>��</div>
            <br>
            <table border=1 cellpadding=2 cellspacing=0 width="95%" align="right">
              <%If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
	CurPage = 1
Else
	CurPage = CINT(Request.QueryString("CurPage"))
End If
sub diaryclose(mess)
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��"& mess &"');location.href = 'list.asp';</script>"
	response.end
end sub
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
select case Request.QueryString("fs")
case "1"
	name=trim(Request.QueryString("name"))
	rs1.Open "Select * From �ռ��û� where ����='"& name &"'", conn1, 1,1
	if rs1.eof and rs1.bof then
		call diaryclose("���û����������⣬�޷��鿴��")
	end if
	bb4=rs1("�ʻ�")
	bb5=rs1("����")
	bb6=rs1("�ռ���")
	bb7=rs1("������")
	bb8=rs1("��������")
	bb9=rs1("�Ĳ�")
	mylb(1)=rs1("lb1")
	mylb(2)=rs1("lb2")
	mylb(3)=rs1("lb3")
	rs1.close
	rs1.Open "Select * From �ռ� where �û���='"& name &"' Order By ��д���� DESC", conn1, 1,1
case "2"
	name=trim(Request.QueryString("name"))
	lb=trim(Request.QueryString("zz"))
	rs1.Open "Select * From �ռ��û� where ����='"& name &"'", conn1, 1,1
	if rs1.eof and rs1.bof then
		call diaryclose("���û����������⣬�޷��鿴��")
	end if
	bb4=rs1("�ʻ�")
	bb5=rs1("����")
	bb6=rs1("�ռ���")
	bb7=rs1("������")
	bb8=rs1("��������")
	bb9=rs1("�Ĳ�")
	mylb(1)=rs1("lb1")
	mylb(2)=rs1("lb2")
	mylb(3)=rs1("lb3")
	rs1.close
	rs1.Open "Select * From �ռ� where �û���='"& name &"' and lb="&lb&" Order By ��д���� DESC", conn1, 1,1
case else
	rs1.Open "Select * From �ռ� where �û���='"& name &"' Order By ��д���� DESC", conn1, 1,1
end select
if rs1.eof and rs1.bof then
%><tr><td align=center >û�м�¼</td></tr></table>
<%
else
rs1.PageSize=15'����ÿҳ��¼��,���޸�          
Dim TotalPages              
TotalPages = rs1.PageCount              
If CurPage>rs1.Pagecount Then               
	CurPage=rs1.Pagecount              
end if                            
rs1.AbsolutePage=CurPage
rs1.CacheSize = rs1.PageSize'��������¼��  
Dim Totalcount
Totalcount =INT(rs1.recordcount)              
StartPageNum=1              
do while StartPageNum+10<=CurPage              
StartPageNum=StartPageNum+10              
Loop              
EndPageNum=StartPageNum+9              
If EndPageNum>rs1.Pagecount then EndPageNum=rs1.Pagecount
I=0
p=rs1.PageSize*(Curpage-1)              
do while (Not rs1.Eof) and (I<rs1.PageSize)              
p=p+1
select case rs1("����")
case "1"
	bm="��������"
case "2"
	bm="��������!"
case "3"
	bm="ƾ����鿴"
case "4"
	bm="["&rs1("��������") &"]�鿴"
case "5"
	bm="["&rs1("��������")&"]������"
case "6"
	bm="["&rs1("��������")&"]�鿴"
end select
%>
              <tr> 
                <td align=left ><img height=14 src="images/hot.gif" width=11><a href="show.asp?id=<%=rs1("id")%>&wj=main.asp" onMouseOver="drs('<div align=center><b>����Ϣ���ϡ�</b></div>�ʻ�:<font color=#ff0000><%=rs1("�ʻ�")%></font>��<br>����:<font color=#ff0000><%=rs1("����")%></font>��<br>����:<font color=#ff0000><%=rs1("������")%></font>��<br>��ʽ:<%=bm%>');return true;" onMouseOut="nd(); return true;"><%=rs1("�ռǱ���")%></a></td>
                <td height=20 width="24%">(<%=rs1("��д����")%>)</td>
              </tr>
              <%
I=I+1
rs1.MoveNext              
Loop
%>
              <tr> 
                <td colspan="2"> 
                  <div align=right> ��<%=Totalcount%>ƪ[ 
                    <% if StartPageNum>1 then %>
                    <a href="list.asp?CurPage=<%=StartPageNum-1%>"><<</a> 
                    <%end if%>
                    <% For I=StartPageNum to EndPageNum
					if I<>CurPage then %>
                    <a href="list.asp?CurPage=<%=I%>"><%=I%></a> 
                    <% else %>
                    <b><font color=#ff0000><%=I%></font></b> 
                    <% end if %>
                    <% Next %>
                    <% if EndPageNum<rs1.Pagecount then %>
                    <a href="list.asp?CurPage=<%=EndPageNum+1%>">>></a> 
                    <%end if%>
                    ] </div>
                </td>
              </tr>
            </table>
            <%end if
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing%>
          </td>
          <td valign=bottom width="5%" rowspan="7"><img src="images/right.gif" width="8" height="289" align="bottom"></td>
        </tr>
        <tr> 
          <td align=left width="1%">&nbsp;</td>
          <td align=left width="20%"> ������<font color=blue><%=rjs%></font>&nbsp;&nbsp;���ԣ�<font color=blue><%=lys%></font></td>
        </tr>
        <tr> 
          <td align=left width="1%">&nbsp;</td>
          <td align=left width="20%"> �ʻ���<font color=blue><%=xhs%></font>&nbsp;&nbsp;������<font color=blue><%=jds%></font></td>
        </tr>
        <tr> 
          <td align=left width="1%">&nbsp;</td>
          <td align=left width="20%"> �Ĳɣ�<font color=blue><%=wcs%></font></td>
        </tr>
        <tr> 
          <td align=left width="1%">&nbsp;</td>
          <td align=left width="20%" valign="top"><a href="add.asp"><img alt=д�����ռ� border=0 height=24 src="images/xie.gif" width=65></a><br>
            <a href="reg.asp?action=edit"><img alt=�޸����� border=0 height=24 src="images/xiugai.gif" width=76></a> 
          </td>
        </tr>
        <tr> 
          <td align=left colspan="2">&nbsp;</td>
        </tr>
        <tr> 
          <td align=left colspan="2" valign="top" height="2"> 
            <div align="left"><br>
            </div>
          </td>
        </tr>
      </table>
    </td>
</tr><tr>
    <td align="right"><img height=44 src="images/down-1.gif" width=574></td>
  </tr></table>
<br>
</body></html>
