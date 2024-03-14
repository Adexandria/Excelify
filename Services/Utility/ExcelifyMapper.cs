using Excelify.Services.Utility.Attributes;
using System.Data;
using System.Reflection;

namespace Excelify.Services.Utility
{
    public class ExcelifyMapper : ExcelMapper
    {
        public override async  Task<List<T>> Map<T>(IEnumerable<DataRow> rows)
        {
            var instances = new List<T>();

            await Task.Run(() =>
            {
                var values = ExcelifyRecord.GetAttribute<ExcelifyAttribute>();
                var properties = typeof(T).GetProperties();
                int left = 0;
                var currentRows = rows.ToList();
                while (left < rows.Count())
                {
                    var instance = (T)Activator.CreateInstance(typeof(T));
                    ExtractProperties<T>(instance, properties, values, currentRows, left);
                    left++;
                    instances.Add(instance);
                }
            });

            return instances;
        }


        private void ExtractProperties<T>(T instance,PropertyInfo[] properties,
            Dictionary<string,object> attributeValues, List<DataRow> currentRows, 
            int rowPosition ,int left = 0)
        {
            if(left <= properties.Length) 
            {
                if (attributeValues.TryGetValue(properties[left].Name, out object? value))
                {
                    var propertyName = value;

                    object propertyValue;
                    if (properties[left].PropertyType == typeof(int) || properties[left].PropertyType == typeof(int?))
                    {
                        propertyValue = ExtractValue<int>(currentRows, propertyName, rowPosition);

                    }
                    else if (properties[left].PropertyType == typeof(Enum))
                    {
                        var currentValue = (string)ExtractValue<string>(currentRows, propertyName, rowPosition);
                        propertyValue = Enum.Parse(properties[left].PropertyType, currentValue);

                    }
                    else if (properties[left].PropertyType == typeof(double) || properties[left].PropertyType == typeof(double?))
                    {
                        propertyValue = ExtractValue<double>(currentRows, propertyName, rowPosition);
                    }
                    else
                    {
                        propertyValue = ExtractValue<string>(currentRows, propertyName, rowPosition);
                    }
                    if (propertyValue != null)
                        properties[left].SetValue(instance, propertyValue);
                }
                ExtractProperties(instance, properties, attributeValues, currentRows ,rowPosition, left++);
            }
        }

        /// <summary>
        /// Extract values based on the field name or field position
        /// </summary>
        /// <typeparam name="TDataType">The data type of the property to map</typeparam>
        /// <param name="rows">Includes the extracted values</param>
        /// <param name="proposedName">The name of the column from the attribute</param>
        /// <param name="left">used for iteration</param>
        /// <returns>an object</returns>
        private object ExtractValue<TDataType>(List<DataRow> rows, object proposedName, int left)
        {
            if (proposedName.GetType().Name == typeof(string).Name)
            {
                return rows[left].Field<TDataType>((string)proposedName);
            }
            else
            {
                return rows[left].Field<TDataType>((int)proposedName);
            }
        }
    }
}
