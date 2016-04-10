using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SrtGenerator.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace SrtGenerator.Serialization
{
    public class XlsParser
    {
        public static SrtModel Load(String path)
        {
            SrtModel mainModel = new SrtModel();

            mainModel.Languages = new List<string>();

            XSSFWorkbook hssfwb;

            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            mainModel.Path = System.IO.Path.GetDirectoryName(path);
            mainModel.Filename = System.IO.Path.GetFileNameWithoutExtension(path);

            ISheet sheet = hssfwb.GetSheetAt(0);

            List<SrtLineModel> models = new List<SrtLineModel>();


            var first_row = sheet.GetRow(0);

            for (int colIndex = 2; colIndex < first_row.Cells.Count; colIndex++)
            {
                mainModel.Languages.Add(first_row.GetCell(colIndex).StringCellValue.Trim());
            }


            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                var currentRow = sheet.GetRow(row);

                if (currentRow.Cells.Count == 0)
                {
                    break;
                }

                var debut = currentRow.GetCell(0);


                if (String.IsNullOrEmpty(debut.StringCellValue))
                {
                    break;
                }

                if (debut.StringCellValue.Trim() == String.Empty)
                {
                    break;
                }


                var fin = currentRow.GetCell(1);

                var language_1_content = currentRow.GetCell(2);

                if (debut == null || fin == null || language_1_content == null)
                {
                    throw new Exception("Ligne n°" + (row + 1) + ": debut ==null ou fin == null ou language_1_content == null");
                }

                DateTime debutTime = DateTime.Now;

                try
                {
                    debutTime = DateTime.ParseExact(debut.StringCellValue.Trim(), "hh:mm:ss,fff", CultureInfo.InvariantCulture);
                }
                catch
                {
                    throw new Exception("Ligne n°" + (row + 1) + ": timecode de debut [" + debut.StringCellValue + "] non valide, le format doit être hh:mm:ss,fff");
                }

                DateTime finTime = DateTime.Now;

                try
                {
                    finTime = DateTime.ParseExact(fin.StringCellValue.Trim(), "hh:mm:ss,fff", CultureInfo.InvariantCulture);
                }
                catch
                {
                    throw new Exception("Ligne n°" + (row + 1) + ": timecode de fin [" + fin.StringCellValue + "] non valide, le format doit être hh:mm:ss,fff");
                }


                Dictionary<String, String> translations = new Dictionary<string, string>();

                for (int colIndex = 2; colIndex < currentRow.Cells.Count; colIndex++)
                {
                    String current_language = mainModel.Languages[colIndex - 2];
                    translations[current_language] = currentRow.GetCell(colIndex).StringCellValue.Trim();
                }


                SrtLineModel model = new SrtLineModel
                {
                    Line = row,
                    StartDate = debutTime,
                    EndDate = finTime,
                    Translations = translations,
                };

                models.Add(model);

            }

            mainModel.SrtLines = models;

            return mainModel;
        }
    }
}
