<!--#include file="Inc/conn.asp"-->
<%
Dim  CurPage
CurPage=cint(Request("page"))
ID=cint(Request("ID"))
If CurPage = empty or CurPage<1 Then 
	 CurPage = 1
End If
Response.ContentType="application/xml"
Response.Charset="gb2312" 
Response.Expires=0
Response.Write("<?xml version=""1.0"" encoding=""gb2312""?>")

Set rs=Server.CreateObject("ADODB.Recordset")
Sql="SELECT * From pcook_Pl where yn = 1 and ArticleID = "&ID&" order by id desc"  
rs.Open Sql,Conn,1,3
IF rs.Eof And rs.Bof Then
    Response.Write("<data page=""0"" P_Nums=""0"">")
	Response.Write("<Content>û������</Content>")
	Response.Write("</data>")
	Response.End
Else
    Dim PerPage,P_Nums,PageCountN,D_Nums
    PerPage=5
	Rs.PageSize=PerPage
	Rs.AbsolutePage=CurPage
    P_Nums=Rs.PageCount 
	D_Nums=Rs.RecordCount
	NoI=0
	Response.Write("<data page="""&CurPage&""" P_Nums="""&P_Nums&""" D_Nums="""&D_Nums&""" ID="""&ID&""">")
    Do Until rs.EOF or PageCountN=PerPage
	NoI=NoI+1
		Response.Write("<NoI>"&rs("ID")&"</NoI>")
		Response.Write("<IP>"&iparray(rs("IP"))&"</IP>")
	  Response.Write("<Author PostTime="""&rs("PostTime")&"""><![CDATA["&left(rs("memAuthor"),5)&"]]></Author>")
	  Response.Write("<Content><![CDATA["&Glhtml(rs("memContent"))&"]]></Content>")
	  Response.Write("<reContent><![CDATA["&rs("reContent")&"]]></reContent>")
	  rs.MoveNext
	  PageCountN=PageCountN+1
	Loop
End If
rs.close
set rs=nothing
Response.Write("</data>")
%>
