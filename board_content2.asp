<html>
<%
db_id = int(request("id"))
 Set db = Server.createObject ("ADODB.Connection")
 dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("board.accdb")
 db.Open (dbMain)

Dim dummy_db, sql
sql = "select * from board_table where ID = " & db_id
Set dummy_db = db.execute(sql)

db_main = dummy_db("db_main")

db_main = Server.HTMLEncode(db_main) 'HTML���ڷ� �ڵ�(��ȯ ���ִ� ��ɾ�) ��, �±������� �����ؼ� �����ش�.
db_main = Replace(db_main, vbLf, vbLf & "<br>") '�ٳ����ֱ�

%>

<body>
<%=db_main%>
</body>
<% db.close %>
</html>
