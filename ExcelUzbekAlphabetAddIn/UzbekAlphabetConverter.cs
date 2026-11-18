namespace ExcelUzbekAlphabetAddIn
{
    internal static class UzbekAlphabetConverter
    {
        internal static string ConvertText(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var result = new System.Text.StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                if (i + 1 < text.Length)
                {
                    string pair = text.Substring(i, 2);
                    string converted = ConvertPair(pair);
                    if (converted != null) { result.Append(converted); i++; continue; }
                }
                result.Append(text[i]);
            }
            return result.ToString();
        }
        private static string ConvertPair(string value)
        {
            switch (value) { case "ch": return "ç"; case "Ch": case "CH": return "Ç"; case "sh": return "ş"; case "Sh": case "SH": return "Ş"; }
            if (value.Length == 2 && (value[1] == '\'' || value[1] == '\u2018' || value[1] == '\u2019' || value[1] == '\u02BB' || value[1] == '\u02BC'))
                switch (value[0]) { case 'o': return "ö"; case 'O': return "Ö"; case 'g': return "ğ"; case 'G': return "Ğ"; }
            return null;
        }
    }
}
