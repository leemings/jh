<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<6 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʬ���ͷ�ֻ�йٸ��ſ��ͷ�');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM �û� where ״̬='��ʬ��' and ���='��ʬ��' order by ���� ",conn
sf=rs("���")
attackname=rs("����")
tl=rs("����")
id=LCase(trim(Request.QueryString("id")))
value=int(abs(clng(Request.QueryString("value"))))
randomize timer
tl=int(rnd*value)+1
select case r
case 0
	yn=0
	Application.Lock
	Application("sjjh_kl")=tl+100000
	Application.UnLock
	abc="<a href='ks.asp?tsk="&Application("sjjh_kl")&"' target='d'><img src='img/cins.GIF' border='0' width='60' height='80'></a>" 
 	attack="<font color=red>���ͷŽ�ʬ����" & sjjh_name & "������ʬ������̫������ע�������������ʬ��Ȼ�������һֻ��ʬ���ͷų�����<bgsound src=wav/gui.wav loop=1>��ʬ��" & attackname & "������"&tl+100000&"<br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>" 
 end select 
if yn=1 then 
	useronlinename=Application("sjjh_useronlinename"&nowinroom) 
	online=Split(Trim(useronlinename)," ",-1) 
	x=UBound(online) 
	Set conn=Server.CreateObject("ADODB.CONNECTION") 
	conn.open Application("sjjh_usermdb") 
	for i=0 to x 

		conn.Execute "update �û� set " & attackname & "=" & attackname & "+" & tl & " where ����='" & online(i) & "'"
	next 
	conn.close 
	set conn=nothing 
end if 
says=attack 
says=replace(says,chr(39),"\'") 
says=replace(says,chr(34),"\"&chr(34)) 
act="��Ϣ" 
towhoway=0 
towho="���" 
addwordcolor="660099" 
saycolor="008888" 
addsays="��" 
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
Response.Redirect "../ok.asp?id=705"
%>

<head>
</head>

</script>















































































































