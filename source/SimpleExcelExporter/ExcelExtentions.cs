using System.Text;
using System.Linq;
namespace SimpleExcelExporter
{
    public static class ExcelExtentions
    {
        public static string SeparateCamelCasingBySpaces(this string value)
        {
            var chars = value.ToCharArray();
            var spacedStringBuilder = new StringBuilder();
            bool isFirstCharacter = true;
            for (int i = 0; i < chars.Count(); i++)
            {
                var currentCharacter = chars[i];

                if (isFirstCharacter)
                {
                    isFirstCharacter = false;

                    spacedStringBuilder.Append(currentCharacter);

                    continue;
                }

                if (char.IsUpper(currentCharacter))
                {
                    spacedStringBuilder.Append(' ');
                }
                spacedStringBuilder.Append(currentCharacter);
            }

            var spacedString = spacedStringBuilder.ToString();

            return spacedString;
        }

        public static object GetValue(this object o, string propertyName)
        {
            var value = o.GetType().GetProperty(propertyName).GetValue(o, null);
            return value;
        }
    }
}