using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using TestClientsProject.Models;

namespace TestClientsProject.CSV
{
    public class ImportCSV
    {
        public static List<Client> GetDataFromCSV(Stream stream)
        {
            var result = new List<Client>();

            using (var reader = new StreamReader(stream))
            {
                bool isFirstRow = true;
                foreach (var row in ReadCsvLine(reader, ','))
                {
                    if (row.Length != 4)
                        break;
                    
                    if(isFirstRow)
                    {
                        isFirstRow = false;
                        continue;
                    }

                    var client = new Client();
                    client.Id = Convert.ToInt32(row[0]);
                    client.Name = row[1];
                    client.Surname = row[2];
                    client.BirthYear = Convert.ToInt32(row[3]);
                    result.Add(client);
                }
            }

            return result;
        }

        private static IEnumerable<string[]> ReadCsvLine(StreamReader reader, char seperator)
        {
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(seperator);

                yield return values;
            }
        }
    }
}