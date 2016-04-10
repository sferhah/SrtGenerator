using SrtGenerator.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace SrtGenerator.Serialization
{
    public class SrtParser
    {
        public static SrtModel Load(String path)
        {
            SrtModel mainModel = new SrtModel();

            mainModel.Languages = new List<string>();

            mainModel.Path = System.IO.Path.GetDirectoryName(path);
            mainModel.Filename = System.IO.Path.GetFileNameWithoutExtension(path);

            List<SrtLineModel> models = new List<SrtLineModel>();
            mainModel.Languages.Add("und");

            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {   
                using (StreamReader reader = new StreamReader(file))
                {
                    SrtLineModel current_model = null;

                    String lineString = null; 

                    while ((lineString = reader.ReadLine())!=null)
                    {   
                        if(lineString=="")
                        {
                            continue;
                        }

                        if(current_model== null)
                        {
                            current_model = new SrtLineModel();
                            current_model.Translations = new Dictionary<string, string>();
                        }

                        if(current_model.Line == 0)
                        {
                            current_model.Line = int.Parse(lineString);
                        }

                        else if(current_model.StartDate == null)
                        {
                            String[] date_array = lineString.Replace(" --> ", "-").Split('-');

                            String debut = date_array[0];
                            String fin = date_array[1];                            

                            try
                            {
                                current_model.StartDate = DateTime.ParseExact(debut, "hh:mm:ss,fff", CultureInfo.InvariantCulture);
                            }
                            catch
                            {
                                throw new Exception("Ligne n°" + (current_model.Line) + ": timecode de debut [" + debut + "] non valide, le format doit être hh:mm:ss,fff");
                            }                            

                            try
                            {
                                current_model.EndDate = DateTime.ParseExact(fin, "hh:mm:ss,fff", CultureInfo.InvariantCulture);
                            }
                            catch
                            {
                                throw new Exception("Ligne n°" + (current_model.Line) + ": timecode de fin [" + fin + "] non valide, le format doit être hh:mm:ss,fff");
                            }
                        }
                        else
                        {
                            current_model.Translations["und"] = lineString;

                            models.Add(current_model);

                            current_model = null;
                        }

                    }
                }
            }


            mainModel.SrtLines = models;

            return mainModel;
        }
    }
}
