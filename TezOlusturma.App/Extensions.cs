namespace TezOlusturma.App
{
    public static class Extensions
    {
        public static string Right(this string value, int length)
        {
            var result = string.IsNullOrEmpty(value) ? "" : value;
            result = result.Length > length ? result.Substring(result.Length - length, length) : result;
            return result;
        }
        public static string Left(this string value, int length)
        {
            var result = string.IsNullOrEmpty(value) ? "" : value;
            result = result.Length > length ? result.Substring(0, length) : result;
            return result;
        }
    }
}
