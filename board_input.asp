<html>

<%
  db_name = request("text_name")
  db_main = request("text_area")
  db_title = request("text_title")
  db_pw = request("text_pw")
  db_date = now()

'스페이스 제거

  For sp_i = 1 to len(db_title)
   if mid(db_title, sp_i, 1) <> chr(32) then
     db_title_x = db_title_x + mid(db_title, sp_i, 1)
   end if
  next

  Set db = Server.createObject ("ADODB.Connection")
  dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("board.accdb")
  db.Open(dbMain)

  insert_db = "insert into board_table(db_title, db_name, db_pw, db_main, db_date, db_lookup, db_spaceno) Values('" & db_title & "','" & db_name & "','" & db_pw & "','" & db_main & "','" & db_date & "','0','" & db_title_x & "')"
  db.execute(insert_db)

  sql = "select * from db_num"
  Set dummy_db = db.execute(sql)

  num_up = int(dummy_db("db_datanum")) + 1

  update_db = "update db_num set db_datanum = '" & num_up & "'"
  db.Execute(update_db)

  db.close()

  %>
  <script language=javascript>
    window.open("board_li.asp", "_self");
  </script>
 </html>
