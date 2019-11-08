<html>
 <%
  Randomize()
  db_co = Int((Rnd * 8) + 1)

  db_name = request("text_name")
  db_main = request("text_main")
  db_date = now()

  Set db = Server.createObject ("ADODB.Connection")
  dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("bang.accdb")
  db.Open(dbMain)

  insert_db = "insert into main_table(db_name, db_main, db_date, db_co) Values('" & db_name & "','" & db_main & "','" & db_date & "','" & db_co & "')"
  db.Execute(insert_db)

  sql = "select * from db_num"
  Set dummy_db = db.execute(sql)

  num_up = int(dummy_db("db_datanum")) + 1

  update_db = "update db_num set db_datanum = '" & num_up & "'"
  db.Execute(update_db)

  db.close()

  %>
  <script language=javascript>
    window.open("guest_book.asp", "_self");
  </script>
 </html>
