using Excelify.Services.Utility.Attributes;


namespace Excelify.Services.Utility
{
    internal class ExcelifyRecord
    {
        /// <summary>
        ///  Get the property name and attribute name
        /// </summary>
        /// <typeparam name="T">Attribute field</typeparam>
        /// <param name="t">Type of the object to read</param>
        /// <returns>A dictionary</returns>
        /// <exception cref="NullReferenceException"></exception>
        public static Dictionary<string, object> GetAttribute<T>() where T : ExcelifyRecordAttribute
        {
            var propertyInfo = typeof(T).GetProperties();
            if (propertyInfo.Length == 0)
                throw new NullReferenceException("properties not found");

            var propertyNames = new Dictionary<string, object>();
            foreach (var info in propertyInfo)
            {
                var attribute = (T)info.GetCustomAttributes(true)[0];
                if (!string.IsNullOrEmpty(attribute.FieldName))
                    propertyNames.Add(info.Name, attribute.FieldName);
                else
                {
                    propertyNames.Add(info.Name, attribute.FieldPosition);
                }
            }
            return propertyNames;
        }
    }
}
