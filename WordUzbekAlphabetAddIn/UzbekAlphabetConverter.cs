namespace WordUzbekAlphabetAddIn
{
    internal static class UzbekAlphabetConverter
    {
        // Every accepted apostrophe form is included because Word keyboards and
        // copied text commonly produce different Unicode apostrophe characters.
        internal static string ConvertPair(string value)
        {
            switch (value)
            {
                case "ch": return "ç";
                case "Ch": return "Ç";
                case "CH": return "Ç";
                case "sh": return "ş";
                case "Sh": return "Ş";
                case "SH": return "Ş";
            }

            if (value.Length == 2 && IsApostrophe(value[1]))
            {
                switch (value[0])
                {
                    case 'o': return "ö";
                    case 'O': return "Ö";
                    case 'g': return "ğ";
                    case 'G': return "Ğ";
                }
            }

            return null;
        }

        private static bool IsApostrophe(char character)
        {
            return character == '\'' || character == '\u2018' || character == '\u2019' ||
                   character == '\u02BB' || character == '\u02BC';
        }
    }
}
