using ExcelDataReader;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace FileProcessor.Services;


public interface IEnumerableFileServices <T>
{
    IEnumerable<T> ParseFile(Stream stream, string filename, bool useHeaderRow, string sheetName = "");
}


public class EnumerableFileServices <T> : IEnumerableFileServices <T>
{

    public EnumerableFileServices ()
    {
    }

    public IEnumerable<T> ParseFile(Stream stream, string filename, bool headerRow = false, string sheetName = "")
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        switch (Path.GetExtension(filename).ToLower())
        {
            case ".xlsx":
                return Parse(stream, headerRow);
            case ".csv":
            case ".txt":
                return ParseCsv(stream, headerRow);
            default:
                throw new ArgumentException("File type not supported");
        }
    }

    

    private IEnumerable<T> Parse(Stream stream, bool headerRow = false)
    {

        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            do
            {
                var h = new Dictionary<int, object>();
                while (reader.Read())
                {
                    if (reader.Depth == 0 && headerRow)
                    {
                        h = Enumerable
                            .Range(0, reader.FieldCount)
                            .ToDictionary(i => i, i => (object)reader.GetValue(i));
                        continue;
                    }

                    var d = new Dictionary<string, object>();
                    d = Enumerable
                        .Range(0, reader.FieldCount)
                        .ToDictionary(i => h[i].ToString()!, i => (object)reader.GetValue(i));
                        
                    var serialized = JsonSerializer.Serialize(d);
                    var deserialized = JsonSerializer.Deserialize<T>(serialized);

                    if (deserialized is null) {
                        throw new Exception($"Unable to deserialize row { serialized }");
                    }

                    yield return deserialized;

                }
                
            } while (reader.NextResult());

        }

        yield break;

    }

    private IEnumerable<T> ParseCsv (Stream stream, bool headerRow = true)
    {
        stream.Seek(0, SeekOrigin.Begin);

        bool first = true;
        
        using (StreamReader sr = new StreamReader(stream))
        {
            var h = new Dictionary<int, object>();
            
            while (!sr.EndOfStream)
            {
                var row = sr.ReadLine();
                if (first && headerRow)
                {
                    var headers = row!.Split(',').ToList();
                    h = Enumerable
                        .Range(0, headers.Count)
                        .ToDictionary(i => i, i => (object)headers[i]);
                    
                    row = sr.ReadLine();
                }

                first = false;

                //https://stackoverflow.com/a/66374354
                var data = Regex.Split(row!, "[,]{1}(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

                var d = new Dictionary<string, object>();
                d = Enumerable
                        .Range(0, h.Count)
                        .ToDictionary(i => h[i].ToString()!, i => (string.IsNullOrWhiteSpace(data[i]) ? null : (object) data[i])!);

                var serialized = JsonSerializer.Serialize(d);
                var deserialized = JsonSerializer.Deserialize<T>(serialized);

                if (deserialized is null) {
                    throw new Exception($"Unable to deserialize row { serialized }");
                }

                yield return deserialized;

            }
        }

        yield break;

    }

}