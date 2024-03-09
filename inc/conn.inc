
<%
dim db
dim weekend
weekend=flase
db="/data/db.mdb"
   dim conn   
   dim connstr

   on error resume next
   connstr="DBQ="+server.mappath(db)+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
   set conn=server.createobject("ADODB.CONNECTION")
   if err then 
      err.clear
   else
        conn.open connstr 
        if err then 
           err.clear
        end if
   end if
%>