<!--#include file="Inc/conn.asp"-->
<!--#include file="Inc/ubb.asp"-->
<!--#include file="Inc.asp"-->
<%
id=CheckStr(Trim(request("id")))
set rs=server.createobject("adodb.recordset")
sql="select * from pcook_Article where id="&id
rs.open sql,conn,1,1
if rs.eof and rs.bof then
Response.write"<script>alert(""����δͨ����� �� ����ȷ��ID"");location.href=""Index.asp"";</script>"
response.end
else

If rs("Linkurl")<>"" then
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ת��<%=rs("title")%></title>
<style>
#ndiv{ 
	width:450px;
	padding:8px;
	margin-top:8px;
	background-color:#FCFFF0;
	border:3px solid #B4EF94;
	height:100px;
	text-align:left;
	font-size:14px;
	line-height:180%;
}
</style>
<meta http-equiv="refresh" content="1;URL=<%=rs("LinkUrl")%>">
</head>
<body>
<div align='center'>
<div id='ndiv'>
����ת��<a href='<%=rs("LinkUrl")%>'><%=rs("LinkUrl")%></a>�����Ժ�...
</div>
</div>
</body>
</html>
<%
Else

set rsClass=server.createobject("adodb.recordset")
sql = "select * from pcook_Class where ID="&rs("ClassID")&""
rsClass.open sql,conn,1,1  
if rsClass.eof and rsClass.bof then
  call Alert("û�д˷���,������ҳ!","index.asp")
  response.end
else
  ClassName=rsClass("ClassName")
rsClass.close
set rsClass=nothing
end if

If rs("PageNum")=0 then
	Content=ManualPagination1(""&rs("ID")&"",""&UBBCode(rs("Content"))&"")
else
	Content=AutoPagination1(""&rs("ID")&"",""&UBBCode(rs("Content"))&"",rs("PageNum"))
End if
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-CN">
<head>
	<title><%=rs("Title")%> - <%=SiteTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
	<meta name="keywords" content="<%=rs("KeyWord")%>" />
	<meta name="description" content="<%=left(LoseHtml(rs("Content")),150)%>" />
	<%if SitePath="/" then%>
	<link href="<%=SitePath%><%=skin%>head.css" rel="stylesheet" type="text/css">
	<%else%>
	<link href="<%=SitePath%><%=skin%>header.css" rel="stylesheet" type="text/css">
	<%end if%>
	<link href="<%=SitePath%><%=skin%>default.css" rel="stylesheet" type="text/css">
	<link href="<%=SitePath%><%=skin%>index.css" rel="stylesheet" type="text/css">
  <script type="text/javascript" src="<%=SitePath%>inc/main.asp"></script>
  <script type="text/javascript" src="<%=SitePath%>inc/main.js"></script>
  <script type="text/javascript" src="<%=SitePath%>Dig.asp"></script>
</head>
<body<%If ""&IsPing&""=1 then%> onLoad="showre(<%=ID%>,1)"<%end if%>>
	<div style="z-index:40000">
	  <%=head%>
		<div class="vtop">
		<div class="vtopbg">
		
		<div class="logo">
		<span class="tlogo"><a href="http://<%=SiteUrl%>" title="<%=SiteTitle%>"></a></span>
		<span class="toptips">������������������������.</span>
		</div>
		

		<div class="menus">
        <span class="tmenu">
        <a href="#">���߽�����</a>
        <a href=# onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('pcook.com.cn');">��Ϊ��ҳ</a>
        <script language="JavaScript">
function bookmarkit()
 {window.external.addFavorite('http://pcook.com.cn','�� Pcook CMS �ݿ� ');}
if (document.all)document.write('<a href="#" onClick="bookmarkit();" title="�ѡ�Pcook CMS �ݿ͡����������ղؼУ�">�ղر�վ</a>')
</script>
        <a target="_blank" href="#" class="end">��վ��ͼ</a>
        </span>
		</div>
		
		
		
		<div class="topmenu">
		
      <div class="menunr">
      <%=Menu%>
      </div>
		
		<div class="topmenubg">
      <span class="lbg"></span>
      <span class="mbg"></span>
      <span class="rbg"></span>
		</div>
		
		</div>

		</div>
	</div>
	<div class="tlinehead"></div>
	
	<div class="vmovie">
		<div class="vmoviebg">
		
<!-- ��࿪ʼ -->

		<div class="vleft">
		

		
<!-- һ��ѭ�����俪ʼ -->

		<div class="movelist">
		<div class="lpic"></div>
		
	
		<div>
		<span class="mbg3">
		<span class="title">�����ڵ�λ�ã�<a href="Index.asp">��ҳ</a> >> <a href="Class.asp?ID=<%=rs("ClassID")%>"><%=ClassName%></a> >> ����</span>
		<span class="more"></span>
		</span>
		</div>
		
		
		<div class="lpic2"></div>
		</div>

		<div class="movies end">
		<h2  style="color:red;font-size:20px;text-align:center;"><%=rs("Title")%></h2>
		<span style="float:right;color:green;">���ߣ�<%=rs("Author")%> ��Դ��<%=rs("CopyFrom")%> ʱ�䣺<%=rs("DateAndTime")%> �����<%=rs("Hits")%></span>
		<BR><BR style="line-height:20px;">
		
		<div class="vmoviesab">
				<p>
				
								<%If ""&ad1&""=1 then%><div id="ad250"><%Call ShowAD(11)%></div><%End if%>
								
								
				<%
				If rs("yn")=2 then 
					If rs("UserName") = pcookuserName then
						Response.Write(""&Content&"")
					else
						Response.Write("<div style=""margin:40px auto;text-align:center;color:#ff0000;"">��Ա˽�����£���ֹ�鿴!</div>") 
					End if
				end if
				
				if rs("yn")=1 then 
				Response.Write("<div style=""margin:40px auto;text-align:center;color:#ff0000;"">�����»�û��ͨ�����</div>") 
				end if
				
				if rs("yn")= 0 then 
				Response.Write(""&Content&"") 
				end if
				%>
				
								<%If ""&ad2&""=1 then%><div id="ad468"><%Call ShowAD(12)%></div><%End if%>
								
      </p>
		</div>

    <div class="vmoviesab">
    <table border="0" cellspacing="0" cellpadding="2" width="80" align="center" style="background:url(<%=SitePath%><%=skin%>dig.gif) no-repeat left 5px;"><tr>
    <td style="width:60px;DISPLAY: block; FONT-WEIGHT: bold; FONT-SIZE: 20px; HEIGHT: 50px; TEXT-ALIGN: center;color:#ff6600;margin-top: 20px;">
		<span id="<%=Rs("id")%>"><%=Rs("dig")%></span>
		</td></tr><tr>
		<td style="width:60px;DISPLAY: block; FONT-WEIGHT: bold; FONT-SIZE: 14px; FONT-FAMILY: Georgia; HEIGHT: 34px; TEXT-ALIGN: center;color:#ff6600;">
		<span id="dig<%=Rs("id")%>"><a href="javascript:Dig(<%=Rs("id")%>);">��һ��</a></span>
		</td></tr></table>	
		</div>

		</div>
		
<!-- һ��ѭ��������� -->
		<hr><br style="line-height:10px;" />
		<div class="movies end">
		<%=thehead%><%=thenext%>
		</div>	

<!-- һ��ѭ�����俪ʼ -->

		<div class="movelist">
		
		<div class="lpic"></div>
		
		<div>
		<span class="mbg3">
		<span class="title">�������</span>
		</span>
		</div>
		
		<div class="lpic2"></div>
		
		</div>

		<div class="movies end">
		
		
		<%=ShowMutualityArticle(""&ID&"",""&rs("KeyWord")&"",8,"��",1)%>
		
		</div>
<!-- һ��ѭ��������� -->
		
	<%If ""&IsPing&""=1 then%>
		<div class="movelist">
		<script type="text/javascript" src="<%=SitePath%>Ajaxpl.asp"></script>
		<div class="lpic"></div>
		<div>
		<span class="mbg3">
		<span class="title">�������</span>
		</span>
		</div>

		<div class="lpic2"></div>
			
			<div id="list"><img src="<%=SiteTable%><%=Skin%>loading.gif" /></div>
			
			<div id="MultiPage"></div>
			
			<div style="clear:both;"></div>
		</div>	
		<div class="movelist">
		<div class="lpic"></div>
		<div>
		<span class="mbg3">
		<span class="title">�����ҵ�����</span>
		</span>
		</div>
		<div class="lpic2"></div>
			<table border="0" cellspacing="0" cellpadding="2" width="600" align="center"><tr>
			
					<td>������<input name="memAuthor" type="text" id="memAuthor" maxlength="8" class="borderall1" value="<%If pcookuserName<>"" then Response.Write(""&pcookuserName&"") else Response.Write("��������") End if%>"/>
					</td><tr><td>
					<li>���ݣ�<textarea name="memContent" cols="30" rows="8" style="width:400px;height:150px;" wrap="virtual" id="memContent" class="borderall1"/></textarea>
					</td><tr><td>
					<input name="ArticleID" type="hidden" id="ArticleID" value="<%=ID%>" />
      <input name="button3" type="button"  class="borderall1" id = "sendGuest" onclick="AddNew()" value="�� ��" />
          </td>
      
	  	</tr></table>

		
		</div>
		
<%end if%>

		</div>
<!-- ������ -->

<!-- �Ҳ࿪ʼ -->

		<div class="vright">
		<div>
		<div  id="upNotice" class="logout" >
		<div class="titles">
		
		<div class="down_con">
		<%Call ShowAD(1)%>	<!-- Pcook ���� -->
		</div>
		
		</div>
		
<!-- ���濪ʼ -->
		<div class="nrlist">
		<div class="bbg2">
		<span class="cib3"></span>
		<span class="mbg2"></span>
		<span class="cib4"></span>
		</div>
		<div id="unlNotice" class="tbg5">
		<%If ""&ad3&""=1 then%><%Call ShowAD(4)%><%End If%>	<!-- Google ��� -->
		</div>
		<div class="bbg2">
		<span class="cib"></span>
		<span class="mbg"></span>
		<span class="cib2"></span>
		</div>
		</div>
<!-- ������� -->
		
		</div>

		</div>
		
		<div class="down_con">
		<%Call ShowAD(5)%>	<!-- ���Ԥ����� -->
		</div>
		
		<div class="hotmovie">
		<span>
		<span class="cit"></span>
		<span class="mbg"></span>
		<span class="cit2"></span>
		</span>
		<span class="tbg">
		<span class="title" style="color:red;">վ���Ƽ�</span>
		<span class="more"><a href="hot.asp">����>></a></span>
		<span class="line"></span>
		
		
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 10 ID,Title,Content,DateAndTime,Hits,TitleFontColor from pcook_Article where yn = 0 and IsHot = 1 order by ID desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
%>
		
		<div class="movies hbg">
		<ul>
		<li class="names"><a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>" title="<%=rs1("title")%>" target="_blank">��<%if rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">"&left(LoseHtml(rs1("title")),30)&"</font>") else Response.Write(""&left(LoseHtml(rs1("title")),30)&"") end if%></a></li>
		</ul>
		</div>
		
<%
  rs1.movenext
  loop
  else
  Response.Write("û��")
  end if
  rs1.close
  set rs1=nothing
%>
		</span>
		<span>
		<div class="cit3"></div>
		<div><span class="mbg2"></span></div>
		<div class="cit4"></div>
		</span>
		</div>

		</div>
		</div>
	</div>
	
	<div class="vbottoms">
	<div class="vbottom">
	<p>
		<%Call ShowAD(6)%>��<%=SiteTcp%>	<!-- ��վ�ײ� -->
	</p>
	<br style="line-height:3px;">
	</div>
</div>
	</div>
<SCRIPT LANGUAGE=JAVASCRIPT><!-- 
if (top.location != self.location)top.location=self.location;
// --></SCRIPT>
<%End if%></body>
<%
rs.close
set rs=nothing
end if

If Request("Page") = "" then 
sql="update pcook_Article set hits = hits + 1 where ID= "&ID
conn.execute(sql)
End if
	
function thehead 
headrs=server.CreateObject("adodb.recordset") 
sql="select top 1 ID,Title from pcook_Article where id<"&id&" and ClassID="&rs("ClassID")&" order by id desc" 
set headrs=conn.execute(sql) 
response.Write("<div class='vmovies'>") 
response.Write("<div class='movernrs'>")
response.Write("<ul>")
if headrs.eof then 
response.Write("<li class='titles'>��һƪ��û����</li>") 
else 
a0=headrs("id") 
a1=headrs("Title")
if seo=1 then 
response.Write("<li class='titles'>��һƪ��<a href='"&SitePath&"Article_"&a0&".html'>"&a1&"</a></li>") 
elseif seo=2 then 
response.Write("<li class='titles'>��һƪ��<a href='"&SitePath&"Article/?"&a0&".html'>"&a1&"</a></li>")  
elseif seo=3 then 
response.Write("<li class='titles'>��һƪ��<a href='"&SitePath&"List.asp?ID="&a0&"'>"&a1&"</a></li>") 
elseif seo=4 then 
response.Write("<li class='titles'>��һƪ��<a href='"&SitePath&"Article/"&a0&".html'>"&a1&"</a></li>") 
end if
end if 
response.Write("</ul>")
response.Write("</div>")
response.Write("</div>")
end function

function thenext 
newrs=server.CreateObject("adodb.recordset") 
sql="select top 1 ID,Title from pcook_Article where id>"&id&" and ClassID="&rs("ClassID")&" order by id desc" 
set newrs=conn.execute(sql)  
response.Write("<div class='vmovies'>") 
response.Write("<div class='movernrs'>")
response.Write("<ul>")
if newrs.eof then 
response.Write("<li class='titles'>��һƪ��û����</li>")
else 
a0=newrs("id") 
a1=newrs("Title")
if seo=1 then 
response.Write("<li class='titles'>��һƪ��<a href='"&SitePath&"Article_"&a0&".html'>"&a1&"</a></li>") 
elseif seo=2 then 
response.Write("<li class='titles'>��һƪ��<a href='"&SitePath&"Article/?"&a0&".html'>"&a1&"</a></li>") 
elseif seo=3 then 
response.Write("<li class='titles'>��һƪ��<a href='"&SitePath&"List.asp?ID="&a0&"'>"&a1&"</a></li>") 
elseif seo=4 then 
response.Write("<li class='titles'>��һƪ��<a href='"&SitePath&"Article/"&a0&".html'>"&a1&"</a></li>") 
end if 
end if 
response.Write("</ul>")
response.Write("</div>")
response.Write("</div>")
end function
%>
</html>