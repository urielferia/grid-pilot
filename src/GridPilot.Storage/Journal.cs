using GridPilot.Core;
using Microsoft.Data.Sqlite;
namespace GridPilot.Storage;

public sealed class Journal : IDisposable
{
    readonly SqliteConnection db;
    public Journal(string path)
    {
        db=new SqliteConnection(new SqliteConnectionStringBuilder{DataSource=path}.ToString());db.Open();
        Execute("PRAGMA journal_mode=WAL; PRAGMA synchronous=FULL; CREATE TABLE IF NOT EXISTS state(key TEXT PRIMARY KEY,value TEXT NOT NULL); CREATE TABLE IF NOT EXISTS operations(id TEXT PRIMARY KEY,plan TEXT NOT NULL,status TEXT NOT NULL,result TEXT);");
    }
    void Execute(string sql) {using var cmd=db.CreateCommand();cmd.CommandText=sql;cmd.ExecuteNonQuery();}
    public void Put<T>(string key,T value) {using var c=db.CreateCommand();c.CommandText="INSERT INTO state VALUES($k,$v) ON CONFLICT(key) DO UPDATE SET value=$v";c.Parameters.AddWithValue("$k",key);c.Parameters.AddWithValue("$v",Json.Write(value));c.ExecuteNonQuery();}
    public T? Get<T>(string key) {using var c=db.CreateCommand();c.CommandText="SELECT value FROM state WHERE key=$k";c.Parameters.AddWithValue("$k",key);return c.ExecuteScalar() is string s && s!="null"?Json.Read<T>(s):default;}
    public (string Plan,string Status,ChangeSet? Result)? Find(string id)
    {
        using var c=db.CreateCommand();c.CommandText="SELECT plan,status,result FROM operations WHERE id=$id";c.Parameters.AddWithValue("$id",id);using var r=c.ExecuteReader();return r.Read()?(r.GetString(0),r.GetString(1),r.IsDBNull(2)?null:Json.Read<ChangeSet>(r.GetString(2))):null;
    }
    public void Begin(ActionPlan plan) {using var c=db.CreateCommand();c.CommandText="INSERT INTO operations(id,plan,status) VALUES($id,$p,'pending')";c.Parameters.AddWithValue("$id",plan.OperationId);c.Parameters.AddWithValue("$p",Json.Hash(Json.Write(plan)));c.ExecuteNonQuery();}
    public void Finish(string id,string status,ChangeSet? result=null) {using var c=db.CreateCommand();c.CommandText="UPDATE operations SET status=$s,result=$r WHERE id=$id";c.Parameters.AddWithValue("$s",status);c.Parameters.AddWithValue("$r",result is null?DBNull.Value:Json.Write(result));c.Parameters.AddWithValue("$id",id);c.ExecuteNonQuery();}
    public void RecoverPending() => Execute("UPDATE operations SET status='interrupted' WHERE status='pending'");
    public void Dispose()=>db.Dispose();
}
