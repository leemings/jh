<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
nowinroom=session("nowinroom")
name=LCase(trim(request.querystring("name")))
aqjh_name=session("aqjh_name")
db=clng(request.querystring("db"))
dbok=clng(request.querystring("aqjh_bingwen"))
if InStr(name,"or")<>0 or InStr(name,"'")<>0 or InStr(name,"`")<>0 or InStr(name,"=")<>0 or InStr(name,"-")<>0 or InStr(name,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
if aqjh_name<>trim(Application("aqjh_dubos")) then
	Response.Write "<script language=JavaScript>{alert('������ʲô���˼�Ҳû��˵����Ĳ���');}</script>"
	Response.End
end if
Application.Lock
	Application("aqjh_dubos")=""
Application.unLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
if db-dbok=-2 or db-dbok=1 then
	Response.Write "<script language=JavaScript>{alert('�����ˣ���ù�ߵģ���~~~');}</script>"
	conn.execute "update �û� set ����=����-100000 where ����='"&aqjh_name&"'"
	conn.execute "update �û� set ����=����+200000 where ����='"&name&"'"
        hunyin="��ù��["& aqjh_name &"]�Ĳ��������{"& name &"}��10�����׻��������ӷ��ˣ�"
end if
if db-dbok=-1 or db-dbok=2 then
        Response.Write "<script language=JavaScript>{alert('��ʤ�ˣ�������10����Ŷ~~~');}</script>"
	conn.execute "update �û� set ����=����+100000 where ����='"&aqjh_name&"'"
	hunyin="��ϲ��["& aqjh_name &"]�Ĳ���Ӯ��{"& name &"}10�������ӣ�"
end if
if db=dbok then
	Response.Write "<script language=JavaScript>{alert('ƽ�ˣ��������ɣ�');}</script>"
    conn.execute "update �û� set ����=����+100000 where ����='"&name&"'"
	hunyin="������["& aqjh_name &"]�Ĳ����{"& name &"}��ƽ�ˣ�ΰ���ͷ�Բ�ı���ϡ�"
end if
'rs.close
'set rs=nothing
conn.close
set conn=nothing
says="<font color=red>��˫�˶Ĳ���</font>"&hunyin		'��������
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
