# xl_service

Программа позволяет экспортировать данные из БД XL Сервисный центр v1.82 (старая версия).

Экспортируются следующие данные:
- id, kv (номер квитанции), username (ФИО клиента), device (устройство), 
- failure (неисправность), acceptdate (дата приёма), readydate (дата готовности),
-  repairstatus (и, необходимый для возможности последующего извещения клиентов (с помощью стороннего сервиса на сайте) статус готовности ремонта)


В программе задействован таймер  
`Timer t = new Timer(tm, null, 0, 900000);`  
Через каждые 15 минут (900000мс) происходит:  
* Подключение к базе данных программы сервисного центра с помощью драйвера Microsoft.Jet.OLEDB  
В строке подключения необходимо изменить Source путь/имя к необходимой базе данных.  
`myConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=база.vdb;Jet OLEDB:System Database=Pattern.mdw;User ID=Excel;Password=lj,thvfy";`  
В строке запроса данные выбираются исходя из струкруты таблиц оригинальной БД программы XL Сервисный центр (https://yadi.sk/d/4Pj0zyF_ZX4DKA)  
`queryString = "SELECT Basa.Kod, Client.Persona,  Phone.Name, Basa.PolomkaDesc, Basa.DataPrihod,  Basa.DataGotov, Basa.KodRepair FROM Basa, Phone, Client WHERE Basa.KodTel = Phone.Kod AND Basa.KodClient = Client.Kod AND Basa.kod>2000";`  
где Basa.kod>2000  — выбор квитанций от номера 2000. Можно установить в 0, чтобы экспротировать все квитанции.

* Подключение к созданной ранее PostgreSQL базе данных, находящейся локально или на сервере,
обновление которой (к сожалению, реализовано только обнуление, т.е. информация в базе данных на время импорта будет отсутствовать)
и является результатом работы этой программы  
`NpgsqlConnection postgreconn = new NpgsqlConnection("Host=localhost;Username=postgres;Password=kurwa;Database=mydb10");`  
localhost или ip сервера, на котором установлен postgresql, логин, пароль и имя базы данных, в которую происходит экспорт.

Более тонкая настройка необходимых полей для экспорта в базу данных требует изменения/добавления части кода.

