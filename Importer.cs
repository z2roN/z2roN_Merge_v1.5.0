using System.Globalization;
using System.Text;
using ExcelDataReader;

namespace ZMerge;

public class ImportResult { public long UploadId; public long Added; public long Skipped; public string Name; public bool CreatedLayout; }

/// 엑셀/CSV 의 첫 시트만 읽어서 누적 저장소에 덧붙인다.
public static class Importer
{
    static string Str(object v) => v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture)?.Trim() ?? "";

    static object Conv(object v)
    {
        switch (v)
        {
            case null: return null;
            case double d: return (d == Math.Floor(d) && Math.Abs(d) < 9e15) ? (object)(long)d : d;
            case DateTime dt: return dt.TimeOfDay == TimeSpan.Zero ? dt.ToString("yyyy-MM-dd") : dt.ToString("yyyy-MM-dd HH:mm:ss");
            case bool b: return b ? "TRUE" : "FALSE";
            case string s: return s.Length == 0 ? null : s;
            default:
                var t = Convert.ToString(v, CultureInfo.InvariantCulture);
                return string.IsNullOrEmpty(t) ? null : t;
        }
    }

    static string[] CleanHeaders(string[] raw)
    {
        var res = new string[raw.Length];
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < raw.Length; i++)
        {
            var h = string.IsNullOrWhiteSpace(raw[i]) ? "열" + (i + 1) : raw[i].Trim();
            if (seen.TryGetValue(h, out var n)) { seen[h] = n + 1; h = h + "_" + (n + 1); } else seen[h] = 1;
            res[i] = h;
        }
        return res;
    }

    /// 들어온 파일의 열 i 가 저장소의 몇 번째 열인지. 열 이름이 같고 순서만 다르면 이름으로 맞춘다.
    static int[] MapColumns(string[] stored, string[] incoming, ISet<string> ignored)
    {
        string Norm(string s) => s.Replace(" ", "").ToLowerInvariant();
        var idx = new Dictionary<string, int>();
        for (int i = 0; i < stored.Length; i++) idx[Norm(stored[i])] = i;
        var map = new int[incoming.Length];
        var extra = new List<string>();
        var used = new HashSet<int>();
        for (int i = 0; i < incoming.Length; i++)
        {
            if (idx.TryGetValue(Norm(incoming[i]), out var k) && used.Add(k)) map[i] = k;
            else if (ignored != null && ignored.Contains(Norm(incoming[i]))) map[i] = -1;   // 사용자가 삭제한 열
            else { map[i] = -1; extra.Add(incoming[i]); }
        }
        var missing = Enumerable.Range(0, stored.Length).Where(i => !used.Contains(i)).Select(i => stored[i]).ToList();
        if (extra.Count == 0 && missing.Count == 0) return map;
        var sb = new StringBuilder("열 구성이 누적 데이터와 다릅니다. 덧붙이지 않았습니다.\n\n");
        sb.AppendLine($"누적 데이터 {stored.Length}개 열 / 올린 파일 {incoming.Length}개 열");
        if (missing.Count > 0) sb.AppendLine("올린 파일에 없는 열: " + string.Join(", ", missing.Take(10)) + (missing.Count > 10 ? " ..." : ""));
        if (extra.Count > 0) sb.AppendLine("누적 데이터에 없는 열: " + string.Join(", ", extra.Take(10)) + (extra.Count > 10 ? " ..." : ""));
        throw new ImportException(sb.ToString());
    }

    public static ImportResult Run(Store st, string file, bool dedupe, int[] keyCols, IProgress<string> prog, CancellationToken ct)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var cfg = new ExcelReaderConfiguration { FallbackEncoding = Encoding.GetEncoding(949) };
        prog?.Report("파일 여는 중: " + Path.GetFileName(file));
        using IExcelDataReader r = file.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ? ExcelReaderFactory.CreateCsvReader(fs, cfg) : ExcelReaderFactory.CreateReader(fs, cfg);
        if (!r.Read()) throw new ImportException("빈 파일입니다.");
        int n = r.FieldCount;
        var raw = new string[n];
        for (int i = 0; i < n; i++) raw[i] = Str(r.GetValue(i));
        // 오른쪽 끝의 완전히 빈 열(헤더 없음)은 제외
        while (n > 0 && raw[n - 1] == "") n--;
        if (n == 0) throw new ImportException("첫 행(헤더)이 비어 있습니다.");
        var headers = CleanHeaders(raw.Take(n).ToArray());

        var res = new ImportResult { Name = Path.GetFileName(file) };
        int[] map;
        bool first = st.Headers.Length == 0;
        if (first) { map = Enumerable.Range(0, n).ToArray(); }
        else map = MapColumns(st.Headers, headers, st.Ignored);
        int cols = first ? n : st.Headers.Length;

        HashSet<string> keys = null;
        int[] kc = null;
        if (dedupe)
        {
            kc = (keyCols == null || keyCols.Length == 0) ? Enumerable.Range(0, cols).ToArray() : keyCols.Where(k => k < cols).ToArray();
            keys = new HashSet<string>();
            if (!first)
            {
                prog?.Report("기존 데이터에서 중복 기준 읽는 중...");
                using var qc = st.Conn.CreateCommand();
                qc.CommandText = $"SELECT {string.Join(",", kc.Select(k => "c" + k))} FROM data";
                using var qr = qc.ExecuteReader();
                var sbk = new StringBuilder();
                while (qr.Read())
                {
                    sbk.Clear();
                    for (int j = 0; j < kc.Length; j++) { if (j > 0) sbk.Append('\u0001'); sbk.Append(qr.IsDBNull(j) ? "" : Convert.ToString(qr.GetValue(j), CultureInfo.InvariantCulture)); }
                    keys.Add(sbk.ToString());
                }
            }
        }

        using var tx = st.Conn.BeginTransaction();
        try
        {
            if (first) { st.CreateTable(headers); res.CreatedLayout = true; }
            long uploadId;
            using (var c = st.Conn.CreateCommand())
            {
                c.Transaction = tx;
                c.CommandText = "INSERT INTO uploads (name,at,rows,skipped) VALUES ($n,$a,0,0)";
                c.Parameters.AddWithValue("$n", res.Name);
                c.Parameters.AddWithValue("$a", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                c.ExecuteNonQuery();
                c.Parameters.Clear();
                c.CommandText = "SELECT last_insert_rowid()";
                uploadId = Convert.ToInt64(c.ExecuteScalar());
            }
            res.UploadId = uploadId;

            using var ins = st.Conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = $"INSERT INTO data (src,{Store.Cols(cols)}) VALUES ($s,{string.Join(",", Enumerable.Range(0, cols).Select(i => "$p" + i))})";
            var pSrc = ins.CreateParameter(); pSrc.ParameterName = "$s"; pSrc.Value = uploadId; ins.Parameters.Add(pSrc);
            for (int i = 0; i < cols; i++) { var p = ins.CreateParameter(); p.ParameterName = "$p" + i; ins.Parameters.Add(p); }

            var vals = new object[cols];
            var sb = new StringBuilder();
            long read = 0;
            while (r.Read())
            {
                if ((++read & 1023) == 0)
                {
                    ct.ThrowIfCancellationRequested();
                    if ((read & 8191) == 0) prog?.Report($"{res.Name}: {read:N0}행 읽는 중...");
                }
                Array.Clear(vals);
                bool any = false;
                int lim = Math.Min(n, r.FieldCount);
                for (int i = 0; i < lim; i++)
                {
                    if (map[i] < 0) continue;
                    var v = Conv(r.GetValue(i));
                    if (v != null) any = true;
                    vals[map[i]] = v;
                }
                if (!any) continue; // 완전히 빈 행은 건너뜀
                if (keys != null)
                {
                    sb.Clear();
                    for (int j = 0; j < kc.Length; j++) { if (j > 0) sb.Append('\u0001'); sb.Append(vals[kc[j]] == null ? "" : Convert.ToString(vals[kc[j]], CultureInfo.InvariantCulture)); }
                    if (!keys.Add(sb.ToString())) { res.Skipped++; continue; }
                }
                for (int i = 0; i < cols; i++) ins.Parameters[i + 1].Value = vals[i] ?? DBNull.Value;
                ins.ExecuteNonQuery();
                res.Added++;
            }
            using (var c = st.Conn.CreateCommand())
            {
                c.Transaction = tx;
                c.CommandText = "UPDATE uploads SET rows=$r, skipped=$k WHERE id=$i";
                c.Parameters.AddWithValue("$r", res.Added); c.Parameters.AddWithValue("$k", res.Skipped); c.Parameters.AddWithValue("$i", uploadId);
                c.ExecuteNonQuery();
            }
            if (res.Added == 0 && first) throw new ImportException("데이터 행이 없습니다. (헤더만 있는 파일)");
            prog?.Report("저장 중...");
            tx.Commit();
        }
        catch
        {
            tx.Rollback();
            if (first) { try { st.ResetAll(); } catch { } }
            throw;
        }
        return res;
    }
}
