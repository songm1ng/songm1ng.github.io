<!--#include file="inc/conn.asp"-->
<% 
dim id 
id=request.querystring("id") 

If Request("Page") = "" then 
sql="update pcook_Article set hits = hits + 1 where ID= "&ID
conn.execute(sql)
End if

	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sql = "Select * From pcook_Article Where id="&id
	Rs.Open Sql,Conn,3,3
'�������ݿ�hit+1����ʡ��..
hits=rs("hits")
'������ʾ���µ�������������ʾ�Ͳ�Ҫ�����
response.write("document.write('�����" & hits & " ��')") 
%>