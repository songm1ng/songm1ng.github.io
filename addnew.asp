<!--#include file="Inc/conn.asp"-->
<%
Dim Author,Content,sResult,ArticleID
Author =CheckStr(unescape(Request.Form("Author")))
Content =CheckStr(unescape(Request.Form("Content")))
ArticleID =CheckStr(unescape(Request.Form("ArticleID")))
IP	=Request.ServerVariables("REMOTE_ADDR")

if Instr(Content,"url=") or Instr(Content,"http://") or Instr(Content,"ad") or Instr(Content,"������") or Instr(Content,"�����") or Instr(Content,"�����") or Instr(Content,"����") or Instr(Content,"����") or Instr(Content,"����") or Instr(Content,"Bë") or Instr(Content,"����") or Instr(Content,"href") or Instr(Content,"����")>0 then
	Response.Write(escape("�벻Ҫ����Υ���������Ϣ!"))
	Response.End
end if

sResult = Author + Content
Conn.Execute("INSERT INTO pcook_Pl(memAuthor,memContent,PostTime,ArticleID,IP,yn)  Values ('"&Author&"','"&Content&"','"&now()&"','"&ArticleID&"','"&IP&"',"&pingoff&")")
If Err Then 
   Response.Write(escape("���ִ���!"))
Else
	If pingoff=0 then
		Response.Write(escape("�������۳ɹ�,������Ҫ����Ա��˺������ʾ��"))
	else
		Response.Write(escape("�������۳ɹ�!"))
	End if
End If
%>