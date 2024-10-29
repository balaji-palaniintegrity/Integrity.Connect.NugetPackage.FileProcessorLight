using ExcelDataReader;
using LightWeightExcelReader;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Text.RegularExpressions;

namespace FileProcessor.Services;

public class FileServiceConfiguration
{
    public int Limit { get; set; }
}

public interface IFileServices
{
    DataSet ParseFile(Stream stream, string fileName, bool useHeaderRow, string sheetName = "");
}
public class FileServices : IFileServices
{
    private readonly FileServiceConfiguration fileConfig;

    public FileServices(
        FileServiceConfiguration fileConfig)
    {
        this.fileConfig = fileConfig;
    }

    public DataSet ParseFile(Stream stream, string filename, bool headerRow = false, string sheetName = "")
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var results = new DataSet();
        switch (Path.GetExtension(filename).ToLower())
        {
            case ".xls":                
            case ".xlsx":
                results = ParseExcel(stream, headerRow, sheetName);
                break;
            //    results = ParseExcelBinaryFileIntoDataSet(stream, headerRow);
            //    break;
            case ".csv":
            case ".txt":
                var csvData = ParseCsv(stream, headerRow);
                results.Tables.Add(csvData);
                break;
            default:
                throw new ArgumentException("File type not supported");
        }
        return results;
    }

    [ExcludeFromCodeCoverage(Justification = "Legacy code using DataSet")]
    private DataSet ParseExcelBinaryFileIntoDataSet(Stream stream, bool headerRow = false)
    {
        using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream))
        {
            var config = new ExcelDataSetConfiguration
            {
                UseColumnDataType = false,
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                {
                    FilterRow = (row) =>
                    {
                        return fileConfig.Limit == 0 || row.Depth <= fileConfig.Limit;
                    },
                    UseHeaderRow = headerRow,
                }
            };
            return excelReader.AsDataSet(config);
        }
    }

    public DataSet ParseExcel(Stream stream, bool useHeaderRow, string sheetName = "")
    {
        var ds = new DataSet();

        var dt = new DataTable();

        var reader = new ExcelReader(stream);

        var sheetReader = reader[sheetName];

        IEnumerable<object> headers = sheetReader.Row(1);

        if (useHeaderRow)
        {
            foreach (var header in headers)
            {
                if (header != null)
                    dt.Columns.Add(GetUniqueColumnName(dt, header.ToString()!));
            }
        }

        bool readFile = true;
        int currentRow = 2;

        while (readFile)
        {                
            AddRowToDataTable(ref dt, sheetReader.Row(currentRow));

            if (fileConfig.Limit > 0 && currentRow == fileConfig.Limit)
            {
                readFile = false;
                break;
            }

            currentRow++;
        }

        ds.Tables.Add(dt);

        return ds;
    }

    private DataTable ParseCsv(Stream stream, bool headerRow = true)
    {
        stream.Seek(0, SeekOrigin.Begin);

        DataTable dt = new DataTable();

        using (StreamReader sr = new StreamReader(stream))
        {
            var headers = new List<string>();

            while (!sr.EndOfStream)
            {
                var row = sr.ReadLine();
                if (dt.Rows.Count == 0 && headerRow)
                {
                    headers = row!.Split(',').ToList();
                    foreach (string header in headers)
                    {
                        dt.Columns.Add(new DataColumn
                        {
                            ColumnName = header,
                            AllowDBNull = true
                        });
                    }
                    row = sr.ReadLine();
                }
                //https://stackoverflow.com/a/66374354
                var data = Regex.Split(row!, "[,]{1}(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

                // May need this if there are more data columns than headers.
                if (dt.Columns.Count < data!.Length)
                {
                    dt.Columns.AddRange(
                        Enumerable.Range(headers.Count, data.Length - headers.Count)
                            .Select(x => new DataColumn($"C{x}"))
                            .ToArray());
                }
                dt.Rows.Add(dt.NewRow()
                    .ItemArray = data!
                        //.Split(',')
                        .Select(x => string.IsNullOrWhiteSpace(x) ? null : TrimQuotes(x))
                        .ToArray());

                if (fileConfig.Limit > 0 && dt.Rows.Count == fileConfig.Limit)
                {
                    break;
                }
            }
        }

        if (dt.Rows.Count == 0)
        {
            throw new ArgumentNullException("Empty file");
        }

        return dt;
    }

    private string GetUniqueColumnName(DataTable table, string name)
    {
        string text = name;
        int num = 1;
        while (table.Columns[text] != null)
        {
            text = $"{name}_{num}";
            num++;
        }

        return text;
    }

    private void AddRowToDataTable(ref DataTable dt, IEnumerable<object> row)
    {
        DataRow dataRow = dt.NewRow();

        for (int i = 0; i < dt.Columns.Count; i++)
        {
            dataRow[i] = row.ElementAt(i);
        }

        dt.Rows.Add(dataRow);
    }

    private string TrimQuotes(string text)
    {
        //ref: https://www.codeproject.com/Questions/5334575/How-do-I-read-a-string-within-double-quotes-from-a
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }
        return text[0] == '"' && text.Length > 2
                   ? text.Substring(1, text.Length - 2).Replace("\"\"", "\"")
                   : text.Replace("\"\"", "\"");
    }
}
