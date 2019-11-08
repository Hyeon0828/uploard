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

db_main = Server.HTMLEncode(db_main) 'HTML문자로 코딩(변환 해주는 명령어) 즉, 태그적용을 해제해서 보여준다.
db_main = Replace(db_main, vbLf, vbLf & "<br>") '줄내려주기

%>

<body>
<%=db_main%>
</body>
<% db.close %>
</html>
