using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using Npgsql;
using System.Threading;
namespace konnekt_xlservice

{
    class Program
    {
        static void Main(string[] args)
        {

            TimerCallback tm = new TimerCallback(Operate);
            Timer t = new Timer(tm, null, 0, 900000);
            Console.ReadLine();

        }




        public static void Operate(object obj)
        {

            OleDbConnection myConn = new OleDbConnection();
            myConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=база.vdb;Jet OLEDB:System Database=Pattern.mdw;User ID=Excel;Password=lj,thvfy";
            string queryString = "SELECT Basa.Kod, Client.Persona,  Phone.Name, Basa.PolomkaDesc, Basa.DataPrihod,  Basa.DataGotov, Basa.KodRepair FROM Basa, Phone, Client WHERE Basa.KodTel = Phone.Kod AND Basa.KodClient = Client.Kod AND Basa.kod>2000";


            ReadMyData(myConn.ConnectionString, queryString);

        }
      





        public static void ReadMyData(string connectionString, string queryString)
        {

            OleDbConnection oleconnection = new OleDbConnection(connectionString);
            OleDbCommand olecommand = new OleDbCommand(queryString, oleconnection);
            oleconnection.Open();
            OleDbDataReader olereader = olecommand.ExecuteReader();


            int i = 1;

            NpgsqlConnection postgreconn = new NpgsqlConnection("Host=localhost;Username=postgres;Password=kurwa;Database=mydb10");
            postgreconn.Open();

            NpgsqlCommand cmddelete = new NpgsqlCommand("DELETE FROM repair *", postgreconn);
            cmddelete.ExecuteNonQuery();


            while (olereader.Read())
            {

                Console.WriteLine(olereader.GetValue(0) +
                                     "  " + olereader.GetValue(1) +
                                     "  " + olereader.GetValue(2) + "  " + olereader.GetValue(3) +
                                     "  " + olereader.GetValue(4) +
                                     "  " + olereader.GetValue(5) +
                                     "  " + olereader.GetValue(6));




                i++;

                

                NpgsqlCommand cmd = new NpgsqlCommand("INSERT INTO repair (id, kv, username, device, failure, acceptdate, readydate, repairstatus) VALUES (@id, @kv, @username, @device, @failure, @acceptdate, @readydate, @repairstatus)", postgreconn);

                // Console.WriteLine(i);

                cmd.Parameters.AddWithValue("id", i);

                //Basa.Kod, Client.Persona,  Phone.Name, Bas.PolomkaDesc, Basa.DataPrihod,  Basa.DataGotov, Basa.KodRepair 

                var val2 = olereader.GetValue(0);
                if (val2 == null) { val2 = 1234554321; }
                cmd.Parameters.AddWithValue("kv", val2);

                // Console.WriteLine(val2);

                var val = olereader.GetValue(1);
                if (val == null || val.ToString().Equals("")) { val = "Информация отсутствует"; }
                cmd.Parameters.AddWithValue("username",   val);

              //  Console.WriteLine(val);

                val = olereader.GetValue(2);
                if (val == null || val.ToString().Equals("")) { val = "Информация отсутствует"; }
                cmd.Parameters.AddWithValue("device",     val);

           //     Console.WriteLine(val);

                val = olereader.GetValue(3);
                if (val == null || val.ToString().Equals("")) { val = "Информация отсутствует"; }
                cmd.Parameters.AddWithValue("failure",    val);

                val = olereader.GetValue(4);
                if (val == null || val.ToString().Equals("")) { val = "Информация отсутствует"; }
                cmd.Parameters.AddWithValue("acceptdate", val);

                val = olereader.GetValue(5);
                if (val == null || val.ToString().Equals("")) { val = "Информация отсутствует"; }
                cmd.Parameters.AddWithValue("readydate", val);


                val2 = olereader.GetValue(6);
                  switch(val2)
                {
                    case 1:
                        val = "В ремонте";
                            break;
                        case 2:
                        val = "Отремонтирован";
                            break;
                        case 0:
                            val = "Без ремонта";
                            break;
                        case null:
                            val = "Информация отсутствует";
                            break;
                    default:
                        val = "Информация отсутствует";
                        break;

                }

                cmd.Parameters.AddWithValue("repairstatus", val);
                cmd.ExecuteNonQuery();


               



            }


            postgreconn.Close();
            oleconnection.Close();
            Console.ReadKey();


        }










    }

}
