using System;
using System.Data;
using Npgsql;
using ETLProcess; 



namespace ETLProcess
{
    internal class DataETL
    {
        private string connectionString = "Host=localhost;Username=postgres;Password=root;Database=Registre";

        public void RunETL()
        {
            var data = ExtractData();
            var transformedData = TransformData(data);
            LoadData(transformedData);
        }

        private DataTable ExtractData()
        {
            DataTable dt = new DataTable();
            using (var conn = new NpgsqlConnection(connectionString))
            {
                conn.Open();
                string sql = @"SELECT a.id_art, a.qte as qte_achats, v.qte as qte_ventes
                               FROM achats a
                               LEFT JOIN ventes v ON a.id_art = v.id_art";
                using (var cmd = new NpgsqlCommand(sql, conn))
                using (var adapter = new NpgsqlDataAdapter(cmd))
                {
                    adapter.Fill(dt);
                }
            }
            return dt;
        }

        private DataTable TransformData(DataTable data)
        {
            DataTable transformedData = new DataTable();
            transformedData.Columns.Add("art_id", typeof(int));
            transformedData.Columns.Add("qte_actuelle", typeof(int));

            foreach (DataRow row in data.Rows)
            {
                int art_id = row.Field<int>("id_art");
                int qte_achats = row.IsNull("qte_achats") ? 0 : row.Field<int>("qte_achats");
                int qte_ventes = row.IsNull("qte_ventes") ? 0 : row.Field<int>("qte_ventes");
                int qte_actuelle = qte_achats - qte_ventes;

                transformedData.Rows.Add(art_id, qte_actuelle);
            }

            return transformedData;
        }

        private void LoadData(DataTable data)
        {
            using (var conn = new NpgsqlConnection(connectionString))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    foreach (DataRow row in data.Rows)
                    {
                        cmd.CommandText = $"INSERT INTO bilan (art_id, qte_actuelle) VALUES (@artId, @qteActuelle) ON CONFLICT (art_id) DO UPDATE SET qte_actuelle = EXCLUDED.qte_actuelle;";
                        cmd.Parameters.AddWithValue("@artId", row["art_id"]);
                        cmd.Parameters.AddWithValue("@qteActuelle", row["qte_actuelle"]);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }
                }
            }
        }
    }
}
