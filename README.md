<div align="center">

## Create your own Grid


</div>

### Description

The code in here is an example of how you can display and manipulate multiple records through asp.
 
### More Info
 
The code works against the Pubs database that comes with SQL Server.

Customizations Required :

Change the Connection String. You may even use the descriptive connection string instead of using UDL file.

Can put the queries in a stored procedure and then call the stored procedure from the web page.

Can add Insert/Delete functionality too.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sumit Dhingra](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sumit-dhingra.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sumit-dhingra-create-your-own-grid__4-6832/archive/master.zip)





### Source Code

```
<%@ Language="VBScript" %>
<%
	option explicit
	Response.Buffer = True
	Dim StrConn
	Dim objConnection
	Dim AuthorsRS, PubsRS
	Dim au_id, Saveflg, Items, Title_Id, stor_Id, Ord_Num, Qty
	Dim SQL
	Dim i, Counter
	'Setting up the Connection
	StrConn = "File Name=C:\Test\Pubs.udl"
	Set objConnection = Server.CreateObject("ADODB.Connection")
	objConnection.ConnectionString = StrConn
	' open the connection
	objConnection.Open
	' Get the Form Values
	SaveFlg = Request.Form("hSaveFlg")
	au_id = Request.Form("hau_id")
	Items = Request.Form("hItems")
	If au_id = "" or ISNULL(au_id) then
		au_id = "%"
	End If
	'---------------------------- Saving the Info to the Database ---------------------------------------
	If SaveFlg = "1" Then
		Counter = 1
		Do While Counter <= (Items - 1)
			If Request.Form("Select"&Counter) = "on" Then
				Ord_Num = Request.Form("Ord_num"&Counter)
				Stor_Id = Request.Form("Stor_Id"&Counter)
				Title_Id = Request.Form("Title_Id"&Counter)
				Qty = Request.Form("Quantity"&Counter)
				' Saving the Info now
				SQL = "Update Sales Set Qty = " & qty & " Where stor_id = " & stor_id &_
				" and ord_num = '" & ord_num & "' and title_id = '" & title_id & "'"
				objConnection.Execute(SQL)
			End If
			Counter = counter + 1
		Loop
	End If
	' Get the Data for displaying on the Page.
	Set AuthorsRS  = objConnection.Execute("Select au_id, au_fname + ' ' + au_lname as Author From Authors")
	AuthorsRS.MoveFirst
	SQL = "select au_fname, au_lname, s.ord_num, t.title, s.qty, s.stor_id, s.title_id " &_
	"from authors a inner join titleauthor ta on a.au_id = ta.au_id inner join sales s " &_
	"on ta.title_id = s.title_id inner join titles t on s.title_id = t.title_id " &_
	"where a.au_id like '" & au_id &	"' order by a.au_fname, au_lname, s.ord_num, s.title_id"
	Set PubsRS  = objConnection.Execute(SQL)
	PubsRS.MoveFirst
%>
<HTML>
<HEAD>
<TITLE>Pubs Grid</TITLE>
<!-- Script to check the fields to see if they are valid -->
<SCRIPT Language="JavaScript">
<!--
function SaveForm()
{
	document.PubsGrid.hSaveFlg.value = 1;
	document.PubsGrid.submit();
}
function SelectAll()
{
	for(i=1; i < document.PubsGrid.hItems.value; i++)
	{
		eval("document.PubsGrid.Select" + i + ".checked = true;");
	}
}
function ClearAll()
{
	for(i=1; i < document.PubsGrid.hItems.value; i++)
	{
		eval("document.PubsGrid.Select" + i + ".checked = false;");
	}
}
function ViewAuthor()
{
	document.PubsGrid.hau_id.value = document.PubsGrid.Author.value;
	document.PubsGrid.submit();
}
-->
</SCRIPT>
</HEAD>
<BODY>
<Form method="POST" name="PubsGrid" action="Grid.asp">
	<Table border="1" cellspacing="0" cellpadding="2">
		<tr>
			<td colspan=4>
				Author :
				<%
					Response.Write "<Select Name=Author onChange='ViewAuthor()'>"
					Response.Write "<option value=All>All</option>"
					Do While Not AuthorsRS.Eof
							Response.Write "<option value=" & AuthorsRS("au_id").Value
							If AuthorsRS("au_id").Value = au_id Then Response.Write " " & "Selected" end If
							Response.Write ">" & AuthorsRS("Author").Value & "</option>"
							AuthorsRS.MoveNext
					Loop
					Response.Write "</Select><P>"
				%>
			</td>
		</tr>
		<tr height="30">
			<td colspan="4"><Input Type=Radio Name=All onClick="return SelectAll()">&nbsp;&nbsp;Select All
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<Input Type=Radio Name=All onClick="return ClearAll()">&nbsp;&nbsp;Clear All
			</td>
		</tr>
		<tr>
			<td>Select</td>
			<td>Order #</td>
			<td>Title</td>
			<td>Quantity</td>
		</tr>
		<%
			i = 1
			Do While Not PubsRS.Eof
				Response.Write "<tr>"
				Response.Write "<td> <Input Type=CheckBox Name=Select" & i & "></td>"
				Response.Write "<td>"
				Response.Write "<Input Type=Hidden Name=Stor_Id" & i & " Value=" & PubsRS("Stor_Id") & ">"
				Response.Write "<Input Type=Hidden Name=Title_Id" & i & " Value=" & PubsRS("title_Id") & ">"
				Response.Write "<Input Type=Text Name=Ord_Num" & i & " Value=" & PubsRS("Ord_num") & ">"
				Response.Write "</td>"
				Response.Write "<td>" & PubsRS("Title") & "</td>"
				Response.Write "<td> <Input Type=Text Name=Quantity" & i & " Value=" & PubsRS("Qty") & "></td>"
				Response.Write "</tr>"
				i = i + 1
				PubsRS.MoveNext
			Loop
		%>
	</Table>
	<input type="Button" value ="Save" name="submit_keyword" onClick="return SaveForm()">
	<input type="Hidden" name="hSaveFlg">
	<input type="Hidden" name="hau_id" value="<%=au_id%>">
	<input type="Hidden" name="hItems" value="<%=i%>">
</Form>
<%
PubsRS.Close
%>
</BODY>
</HTML>
```

