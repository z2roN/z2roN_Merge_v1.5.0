using System.Globalization;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using MiniExcelLibs;

namespace ZMerge;

public class ImportException : Exception { public ImportException(string m) : base(m) { } }

public class UploadInfo { public long Id; public string Name; public string At; public long Rows; public long Skipped; }

/// 누적 데이터 저장소 (SQLite 1개 파일). 열은 c0, c1, ... 로 저장하고 헤더 이름은 meta 에 보관한다.
public class Store : IDisposable
{
    public const int Block = 500;
    public const int ExcelMaxRows = 1048575; // 헤더 1행 제외
    SqliteConnection conn;
    public string DbPath { get; private set; }
    public string[] Headers { get; private set; } = Array.Empty<string>();
    public SqliteConnection Conn => conn;
    /// 사용자가 삭제한 열 이름(정규화). 이후 올리는 파일에 이 열이 있어도 무시한다.
    public HashSet<string> Ignored { get; private set; } = new();
    public static string NormName(string s) => s.Replace(" ", "").ToLowerInvariant();

    public static Store Open(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        var s = new Store { DbPath = path };
        s.conn = new SqliteConnection("Data Source=" + path);
        s.conn.Open();
        s.Exec("PRAGMA journal_mode=WAL");
        s.Exec("PRAGMA synchronous=NORMAL");
        s.Exec("CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)");
        s.Exec("CREATE TABLE IF NOT EXISTS uploads (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, at TEXT, rows INTEGER, skipped INTEGER)");
        var h = s.Scalar("SELECT v FROM meta WHERE k='headers'") as string;
        if (!string.IsNullOrEmpty(h)) s.Headers = JsonSerializer.Deserialize<string[]>(h);
        var ig = s.Scalar("SELECT v FROM meta WHERE k='ignored'") as string;
        if (!string.IsNullOrEmpty(ig)) s.Ignored = new HashSet<string>(JsonSerializer.Deserialize<string[]>(ig));
        return s;
    }

    public void Dispose()
    {
        try { if (conn != null && conn.State == System.Data.ConnectionState.Open) Exec("PRAGMA wal_checkpoint(TRUNCATE)"); } catch { }
        try { conn?.Close(); conn?.Dispose(); } catch { }
        conn = null;
        SqliteConnection.ClearAllPools();
    }

    public int Exec(string sql, params (string, object)[] ps)
    {
        using var c = conn.CreateCommand();
        c.CommandText = sql;
        foreach (var (k, v) in ps) c.Parameters.AddWithValue(k, v ?? DBNull.Value);
        return c.ExecuteNonQuery();
    }

    public object Scalar(string sql, params (string, object)[] ps)
    {
        using var c = conn.CreateCommand();
        c.CommandText = sql;
        foreach (var (k, v) in ps) c.Parameters.AddWithValue(k, v ?? DBNull.Value);
        var r = c.ExecuteScalar();
        return r is DBNull ? null : r;
    }

    public static string Cols(int n) => string.Join(",", Enumerable.Range(0, n).Select(i => "c" + i));

    public void CreateTable(string[] headers)
    {
        Exec($"CREATE TABLE IF NOT EXISTS data (rid INTEGER PRIMARY KEY AUTOINCREMENT, src INTEGER, {Cols(headers.Length)})");
        Exec("CREATE INDEX IF NOT EXISTS ix_src ON data(src)");
        Headers = headers;
        Exec("INSERT OR REPLACE INTO meta (k,v) VALUES ('headers',$v)", ("$v", JsonSerializer.Serialize(headers)));
    }

    /// 열 순서 변경/삭제. keep[새 위치] = 기존 열 번호. keep 에 없는 열은 삭제된다. (삭제 취소 기록은 비워짐)
    public void RestructureColumns(int[] keep)
    {
        int n = Headers.Length;
        if (keep.Length == 0) throw new InvalidOperationException("열을 모두 삭제할 수는 없습니다.");
        if (keep.Distinct().Count() != keep.Length || keep.Any(k => k < 0 || k >= n)) throw new ArgumentException("잘못된 열 목록");
        var drop = Enumerable.Range(0, n).Where(i => !keep.Contains(i)).ToArray();
        Exec("DROP TABLE IF EXISTS dupview");
        Exec("DROP TABLE IF EXISTS trash");
        using (var tx = conn.BeginTransaction())
        {
            void Run(string sql) { using var c = conn.CreateCommand(); c.Transaction = tx; c.CommandText = sql; c.ExecuteNonQuery(); }
            foreach (var d in drop) Run($"ALTER TABLE data DROP COLUMN c{d}");
            for (int p = 0; p < keep.Length; p++) if (keep[p] != p) Run($"ALTER TABLE data RENAME COLUMN c{keep[p]} TO k{p}");
            for (int p = 0; p < keep.Length; p++) if (keep[p] != p) Run($"ALTER TABLE data RENAME COLUMN k{p} TO c{p}");
            var newHeaders = keep.Select(k => Headers[k]).ToArray();
            var ign = new HashSet<string>(Ignored);
            foreach (var d in drop) ign.Add(NormName(Headers[d]));
            foreach (var h in newHeaders) ign.Remove(NormName(h));
            using (var c = conn.CreateCommand()) { c.Transaction = tx; c.CommandText = "INSERT OR REPLACE INTO meta (k,v) VALUES ('headers',$v)"; c.Parameters.AddWithValue("$v", JsonSerializer.Serialize(newHeaders)); c.ExecuteNonQuery(); }
            using (var c = conn.CreateCommand()) { c.Transaction = tx; c.CommandText = "INSERT OR REPLACE INTO meta (k,v) VALUES ('ignored',$v)"; c.Parameters.AddWithValue("$v", JsonSerializer.Serialize(ign.ToArray())); c.ExecuteNonQuery(); }
            tx.Commit();
            Headers = newHeaders; Ignored = ign;
        }
    }

    public void ResetAll()
    {
        Exec("DROP TABLE IF EXISTS data");
        Exec("DROP TABLE IF EXISTS trash");
        Exec("DELETE FROM meta");
        Ignored = new();
        Exec("DELETE FROM uploads");
        Headers = Array.Empty<string>();
        Exec("VACUUM");
    }

    public static string KeyOrder(int[] cols) => string.Join(",", cols.Select(k => $"c{k} COLLATE NOCASE ASC")) + ", rid ASC";

    /// 선택한 열 값이 같은 행(2개 이상)을 모두 임시 테이블 dupview 에 담는다. 화면에서 중복 행만 보여주는 데 사용.
    public long BuildDupView(int[] cols, bool ignoreBlank, bool normalize, out long groups)
    {
        Exec("DROP TABLE IF EXISTS temp.dupview");
        Exec("CREATE TEMP TABLE dupview (rid INTEGER PRIMARY KEY)");
        var first = new Dictionary<string, long>();
        var dupRids = new HashSet<long>();
        var dupKeys = new HashSet<string>();
        using (var c = conn.CreateCommand())
        {
            c.CommandText = $"SELECT rid,{string.Join(",", cols.Select(k => "c" + k))} FROM data";
            using var r = c.ExecuteReader();
            var sb = new StringBuilder();
            while (r.Read())
            {
                sb.Clear();
                bool allBlank = true;
                for (int j = 0; j < cols.Length; j++)
                {
                    var v = r.IsDBNull(j + 1) ? "" : Convert.ToString(r.GetValue(j + 1), CultureInfo.InvariantCulture) ?? "";
                    if (normalize) v = v.Trim().ToLowerInvariant();
                    if (v != "") allBlank = false;
                    sb.Append(v).Append('\u0001');
                }
                if (ignoreBlank && allBlank) continue;
                var key = sb.ToString();
                long rid = r.GetInt64(0);
                if (first.TryGetValue(key, out var f0)) { dupRids.Add(rid); dupRids.Add(f0); dupKeys.Add(key); }
                else first[key] = rid;
            }
        }
        groups = dupKeys.Count;
        using var tx = conn.BeginTransaction();
        using (var ins = conn.CreateCommand())
        {
            ins.Transaction = tx;
            ins.CommandText = "INSERT INTO dupview (rid) VALUES ($r)";
            var p = ins.CreateParameter(); p.ParameterName = "$r"; ins.Parameters.Add(p);
            foreach (var id in dupRids) { p.Value = id; ins.ExecuteNonQuery(); }
        }
        tx.Commit();
        return dupRids.Count;
    }

    /// 선택한 열 값이 같은 행 중, order 기준 첫 행만 남기고 나머지의 rid 를 돌려준다.
    public List<long> FindDuplicateRids(int[] cols, string order, bool ignoreBlank, bool normalize, out long groups)
    {
        var seen = new HashSet<string>();
        var dups = new List<long>();
        var dupKeys = new HashSet<string>();
        using var c = conn.CreateCommand();
        c.CommandText = $"SELECT rid,{string.Join(",", cols.Select(k => "c" + k))} FROM data ORDER BY {order}";
        using var r = c.ExecuteReader();
        var sb = new StringBuilder();
        while (r.Read())
        {
            sb.Clear();
            bool allBlank = true;
            for (int j = 0; j < cols.Length; j++)
            {
                var v = r.IsDBNull(j + 1) ? "" : Convert.ToString(r.GetValue(j + 1), CultureInfo.InvariantCulture) ?? "";
                if (normalize) v = v.Trim().ToLowerInvariant();
                if (v != "") allBlank = false;
                sb.Append(v).Append('\u0001');
            }
            if (ignoreBlank && allBlank) continue;
            var key = sb.ToString();
            if (!seen.Add(key)) { dups.Add(r.GetInt64(0)); dupKeys.Add(key); }
        }
        groups = dupKeys.Count;
        return dups;
    }

    // ---------- 행 삭제 / 삭제 취소 / 정렬 확정 ----------
    void EnsureTrash() => Exec($"CREATE TABLE IF NOT EXISTS trash (batch INTEGER, rid INTEGER, src INTEGER, {Cols(Headers.Length)})");

    /// 선택한 행(rid)을 삭제한다. 최근 5번의 삭제는 휴지통에 보관되어 되돌릴 수 있다.
    public long DeleteRows(List<long> rids)
    {
        if (rids.Count == 0) return 0;
        EnsureTrash();
        var cols = Cols(Headers.Length);
        using var tx = conn.BeginTransaction();
        long batch;
        using (var c = conn.CreateCommand()) { c.Transaction = tx; c.CommandText = "SELECT COALESCE(MAX(batch),0)+1 FROM trash"; batch = Convert.ToInt64(c.ExecuteScalar()); }
        for (int i = 0; i < rids.Count; i += 5000)
        {
            var list = string.Join(",", rids.GetRange(i, Math.Min(5000, rids.Count - i)));
            using var c1 = conn.CreateCommand(); c1.Transaction = tx;
            c1.CommandText = $"INSERT INTO trash (batch,rid,src,{cols}) SELECT {batch},rid,src,{cols} FROM data WHERE rid IN ({list})"; c1.ExecuteNonQuery();
            using var c2 = conn.CreateCommand(); c2.Transaction = tx;
            c2.CommandText = $"DELETE FROM data WHERE rid IN ({list})"; c2.ExecuteNonQuery();
        }
        using (var c = conn.CreateCommand()) { c.Transaction = tx; c.CommandText = $"DELETE FROM trash WHERE batch <= {batch - 5}"; c.ExecuteNonQuery(); }
        tx.Commit();
        return rids.Count;
    }

    public bool CanUndoDelete() => Headers.Length > 0 && Convert.ToInt64(Scalar("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='trash'")) > 0 && Scalar("SELECT MAX(batch) FROM trash") != null;

    public long UndoDelete()
    {
        if (!CanUndoDelete()) return 0;
        var batch = Convert.ToInt64(Scalar("SELECT MAX(batch) FROM trash"));
        var cols = Cols(Headers.Length);
        using var tx = conn.BeginTransaction();
        long n;
        using (var c = conn.CreateCommand()) { c.Transaction = tx; c.CommandText = $"INSERT INTO data (rid,src,{cols}) SELECT rid,src,{cols} FROM trash WHERE batch={batch}"; n = c.ExecuteNonQuery(); }
        using (var c = conn.CreateCommand()) { c.Transaction = tx; c.CommandText = $"DELETE FROM trash WHERE batch={batch}"; c.ExecuteNonQuery(); }
        tx.Commit();
        return n;
    }

    /// 현재 정렬 순서를 데이터 자체의 순서로 확정(저장)한다. 이후 새로 올리는 행은 맨 아래에 붙는다.
    public void CommitSort(string order)
    {
        var cols = Cols(Headers.Length);
        using var tx = conn.BeginTransaction();
        foreach (var sql in new[]
        {
            "DROP TABLE IF EXISTS trash",
            "ALTER TABLE data RENAME TO data_old",
            $"CREATE TABLE data (rid INTEGER PRIMARY KEY AUTOINCREMENT, src INTEGER, {cols})",
            $"INSERT INTO data (src,{cols}) SELECT src,{cols} FROM data_old ORDER BY {order}",
            "DROP TABLE data_old",
            "CREATE INDEX IF NOT EXISTS ix_src ON data(src)",
        })
        { using var c = conn.CreateCommand(); c.Transaction = tx; c.CommandText = sql; c.ExecuteNonQuery(); }
        tx.Commit();
    }

    public long TotalRows() => Headers.Length == 0 ? 0 : Convert.ToInt64(Scalar("SELECT COUNT(*) FROM data"));

    public List<UploadInfo> Uploads()
    {
        var l = new List<UploadInfo>();
        using var c = conn.CreateCommand();
        c.CommandText = "SELECT id,name,at,rows,skipped FROM uploads ORDER BY id";
        using var r = c.ExecuteReader();
        while (r.Read()) l.Add(new UploadInfo { Id = r.GetInt64(0), Name = r.GetString(1), At = r.GetString(2), Rows = r.GetInt64(3), Skipped = r.GetInt64(4) });
        return l;
    }

    /// 업로드 단위로 되돌리기(삭제). 데이터가 전부 없어지면 열 구성도 초기화해서 새 형식으로 시작할 수 있다.
    public void DeleteUploads(IEnumerable<long> ids)
    {
        var list = string.Join(",", ids);
        if (list == "") return;
        using var tx = conn.BeginTransaction();
        using (var c = conn.CreateCommand()) { c.Transaction = tx; c.CommandText = $"DELETE FROM data WHERE src IN ({list})"; c.ExecuteNonQuery(); }
        using (var c = conn.CreateCommand()) { c.Transaction = tx; c.CommandText = $"DELETE FROM uploads WHERE id IN ({list})"; c.ExecuteNonQuery(); }
        tx.Commit();
        if (TotalRows() == 0) ResetAll();
    }

    // ---------- 조회 ----------
    public static string BuildWhere(string text, int col, long? srcFrom, long? srcTo, int colCount, out List<(string, object)> ps, bool dupOnly = false)
    {
        ps = new List<(string, object)>();
        var parts = new List<string>();
        if (!string.IsNullOrEmpty(text))
        {
            var like = "%" + text.Replace("\\", "\\\\").Replace("%", "\\%").Replace("_", "\\_") + "%";
            ps.Add(("$f", like));
            parts.Add(col >= 0 ? $"c{col} LIKE $f ESCAPE '\\'" : "(" + string.Join(" OR ", Enumerable.Range(0, colCount).Select(i => $"c{i} LIKE $f ESCAPE '\\'")) + ")");
        }
        if (srcFrom != null) { ps.Add(("$s1", srcFrom.Value)); ps.Add(("$s2", srcTo ?? srcFrom.Value)); parts.Add("src BETWEEN $s1 AND $s2"); }
        if (dupOnly) parts.Add("rid IN (SELECT rid FROM dupview)");
        return parts.Count == 0 ? "" : "WHERE " + string.Join(" AND ", parts);
    }

    public static string BuildOrder(List<(int col, bool desc)> sorts, bool reverse = false)
    {
        var o = sorts.Select(s => $"c{s.col} COLLATE NOCASE {(s.desc != reverse ? "DESC" : "ASC")}").ToList();
        o.Add(reverse ? "rid DESC" : "rid ASC");
        return string.Join(", ", o);
    }

    public long Count(string where, List<(string, object)> ps) => Convert.ToInt64(Scalar($"SELECT COUNT(*) FROM data {where}", ps.ToArray()));

    public List<long> QueryIds(string where, List<(string, object)> ps, string order)
    {
        var ids = new List<long>();
        if (Headers.Length == 0) return ids;
        using var c = conn.CreateCommand();
        c.CommandText = $"SELECT rid FROM data {where} ORDER BY {order}";
        foreach (var (k, v) in ps) c.Parameters.AddWithValue(k, v);
        using var r = c.ExecuteReader();
        while (r.Read()) ids.Add(r.GetInt64(0));
        return ids;
    }

    public object[][] Fetch(List<long> ids, int start, int count)
    {
        int end = Math.Min(ids.Count, start + count);
        var slice = ids.GetRange(start, Math.Max(0, end - start));
        int n = Headers.Length;
        var map = new Dictionary<long, object[]>();
        if (slice.Count > 0)
        {
            using var c = conn.CreateCommand();
            c.CommandText = $"SELECT rid,{Cols(n)} FROM data WHERE rid IN ({string.Join(",", slice)})";
            using var r = c.ExecuteReader();
            while (r.Read())
            {
                var row = new object[n];
                for (int i = 0; i < n; i++) row[i] = r.IsDBNull(i + 1) ? null : r.GetValue(i + 1);
                map[r.GetInt64(0)] = row;
            }
        }
        var res = new object[slice.Count][];
        for (int i = 0; i < slice.Count; i++) res[i] = map.TryGetValue(slice[i], out var x) ? x : new object[n];
        return res;
    }

    // ---------- 내보내기 ----------
    string SelectForExport(string where, string order) =>
        "SELECT " + string.Join(",", Headers.Select((h, i) => $"c{i} AS \"{h.Replace("\"", "\"\"")}\"")) + $" FROM data {where} ORDER BY {order} LIMIT $lim OFFSET $off";

    public List<string> ExportXlsx(string path, string where, List<(string, object)> ps, string order, IProgress<string> prog)
    {
        long total = Count(where, ps);
        int parts = (int)Math.Max(1, (total + ExcelMaxRows - 1) / ExcelMaxRows);
        var files = new List<string>();
        for (int p = 0; p < parts; p++)
        {
            var f = parts == 1 ? path : Path.Combine(Path.GetDirectoryName(path)!, Path.GetFileNameWithoutExtension(path) + $"_{p + 1}.xlsx");
            prog?.Report($"엑셀 저장 중... ({p + 1}/{parts})");
            using var c = conn.CreateCommand();
            c.CommandText = SelectForExport(where, order);
            foreach (var (k, v) in ps) c.Parameters.AddWithValue(k, v);
            c.Parameters.AddWithValue("$lim", ExcelMaxRows);
            c.Parameters.AddWithValue("$off", (long)p * ExcelMaxRows);
            using var rd = c.ExecuteReader();
            MiniExcel.SaveAs(f, rd, printHeader: true, sheetName: "합본", overwriteFile: true);
            files.Add(f);
        }
        return files;
    }

    public long ExportCsv(string path, string where, List<(string, object)> ps, string order, IProgress<string> prog)
    {
        using var sw = new StreamWriter(path, false, new UTF8Encoding(true));
        sw.WriteLine(string.Join(",", Headers.Select(Csv)));
        using var c = conn.CreateCommand();
        c.CommandText = $"SELECT {Cols(Headers.Length)} FROM data {where} ORDER BY {order}";
        foreach (var (k, v) in ps) c.Parameters.AddWithValue(k, v);
        using var r = c.ExecuteReader();
        long n = 0;
        var sb = new StringBuilder();
        while (r.Read())
        {
            sb.Clear();
            for (int i = 0; i < Headers.Length; i++)
            {
                if (i > 0) sb.Append(',');
                if (!r.IsDBNull(i)) sb.Append(Csv(Convert.ToString(r.GetValue(i), CultureInfo.InvariantCulture)));
            }
            sw.WriteLine(sb.ToString());
            if (++n % 20000 == 0) prog?.Report($"CSV 저장 중... {n:N0}행");
        }
        return n;
    }

    static string Csv(string s) => s != null && (s.Contains(',') || s.Contains('"') || s.Contains('\n') || s.Contains('\r')) ? "\"" + s.Replace("\"", "\"\"") + "\"" : s ?? "";
}
