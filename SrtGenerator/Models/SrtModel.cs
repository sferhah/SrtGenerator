using System;
using System.Collections.Generic;
using System.Text;

namespace SrtGenerator.Models
{

    public class SrtModel
    {

        public String Path { get; set; }
        public String Filename { get; set; }

        public List<SrtLineModel> SrtLines { get; set; }
        public List<String> Languages { get; set; }

        public String GetSrtText(String language)
        {
            StringBuilder builder = new StringBuilder();

            foreach (var m in SrtLines)
            {
                builder.AppendLine(m.Line.ToString());
                builder.Append(m.StartDate.Value.ToString("HH:mm:ss,fff"));
                builder.Append(" --> ");
                builder.Append(m.EndDate.Value.ToString("HH:mm:ss,fff"));
                builder.AppendLine();

                String value;
                m.Translations.TryGetValue(language, out value);
                builder.AppendLine(value);
                builder.AppendLine();
            }

            String result = builder.ToString();

            return result;
        }
    }
}
