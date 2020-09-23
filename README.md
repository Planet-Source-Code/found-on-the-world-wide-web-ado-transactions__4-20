<div align="center">

## ADO Transactions


</div>

### Description

Transactions are atomic operations that allow you to do multiple operations on a database as one operation. For example, if you were creating a banking application in which you deducted $100 from one account and added it to another account, you wouldn't want the operation to fail right in the middle, because the money would be 'lost'! The solution is to wrap the SQL in a transaction. If the operation is aborted in the middle (the pc gets shut off for example) the database will rollback the changes so that the initial account was never debited the $100. This will make you feel good, especially if its your bank account!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-ado-transactions__4-20/archive/master.zip)





### Source Code

```
<% Response.Expires = 0 %>
<HTML>
<HEAD><TITLE><H3>Transactions</H3></TITLE></HEAD>
<BODY BGColor=ffffff Text=000000>
<%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("guestDSN")
Set rs = Server.CreateObject("ADODB.RecordSet")
MySQL = "SELECT * FROM paulen"
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.ActiveConnection = cn
rs.Source = MySQL
rs.Open
%>
<h2>Before:<BR>
<TABLE BORDER=1>
<TR>
<% For i = 0 to RS.Fields.Count - 1 %>
<TD><B><% = RS(i).Name %></B></TD>
<% Next %>
</TR>
<% Do While Not RS.EOF %>
<TR>
<% For i = 0 to RS.Fields.Count - 1 %>
<TD VALIGN=TOP><% = RS(i) %></TD>
<% Next %>
</TR>
<%
RS.MoveNext
Loop
RS.Close
%>
</TABLE>
<%
cn.BeginTrans
cn.Execute("INSERT INTO paulen (fld1, fld2) VALUES ('Aborted', 50)")
cn.RollbackTrans
cn.BeginTrans
cn.Execute("INSERT INTO paulen (fld1, fld2) VALUES ('Trans" & Time() & "', 100)")
cn.CommitTrans
%>
Completed.<P>
<h2>After:</h2>
<TABLE BORDER=1>
<TR>
<%
rs.Open
For i = 0 to RS.Fields.Count - 1 %>
<TD><B><% = RS(i).Name %></B></TD>
<% Next %>
</TR>
<% Do While Not RS.EOF %>
<TR>
<% For i = 0 to RS.Fields.Count - 1 %>
<TD VALIGN=TOP><% = RS(i) %></TD>
<% Next %>
</TR>
<%
RS.MoveNext
Loop
RS.Close
Cn.Close
%>
</TABLE>
```

