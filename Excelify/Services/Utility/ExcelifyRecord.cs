﻿using Excelify.Services.Utility.Attributes;
using System.ComponentModel;


namespace Excelify.Services.Utility
{
    internal static class ExcelifyRecord
    {
        public static string GetDescription<TAttribute>(this Enum value) where TAttribute : DescriptionAttribute
        {
            var propertyInfo = value.GetType().GetMember(value.ToString());
            if(propertyInfo.Length == 0)
            {
                 throw new NullReferenceException("properties not found");
            }
            if (propertyInfo[0].GetCustomAttributes(false).FirstOrDefault() is DescriptionAttribute attribute)
            {
                return attribute.Description;
            }
            else
            {
                return value.ToString();
            }
        }

        /// <summary>
        ///  Get the property name and attribute name
        /// </summary>
        /// <typeparam name="TAttribute">Attribute field</typeparam>
        /// <param name="t">Type of the object to read</param>
        /// <returns>A dictionary</returns>
        /// <exception cref="NullReferenceException"></exception>
        public static Dictionary<string, object> GetAttribute<TAttribute>(Type t) where TAttribute : ExcelifyRecordAttribute
        {
            var propertyInfo = t.GetProperties();
            if (propertyInfo.Length == 0)
                throw new NullReferenceException("properties not found");

            var propertyNames = new Dictionary<string, object>();
            foreach (var info in propertyInfo)
            {
                if (info.GetCustomAttributes(true).FirstOrDefault() is not ExcelifyRecordAttribute attribute)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(attribute.FieldName))
                    propertyNames.Add(info.Name, attribute.FieldName);
                else
                {
                    propertyNames.Add(info.Name, attribute.FieldPosition);
                }
            }
            return propertyNames;
        }

        public static Dictionary<string, object> GetAttribute<TAttribute,TEntity>() 
            where TAttribute : ExcelifyRecordAttribute
            where TEntity : class
        {
            var propertyInfo = typeof(TEntity).GetProperties();
            if (propertyInfo.Length == 0)
                throw new NullReferenceException("properties not found");

            var propertyNames = new Dictionary<string, object>();
            foreach (var info in propertyInfo)
            {
                if (info.GetCustomAttributes(true).FirstOrDefault() is not ExcelifyRecordAttribute attribute)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(attribute.FieldName))
                    propertyNames.Add(info.Name, attribute.FieldName);
                else
                {
                    propertyNames.Add(info.Name, attribute.FieldPosition);
                }
            }
            return propertyNames;
        }

        public static Dictionary<string,object> GetAttribute(Type attributeType, Type entityType) 
        {
            var propertyInfo = entityType.GetProperties();
            if (propertyInfo.Length == 0)
                throw new NullReferenceException("properties not found");

            var propertyNames = new Dictionary<string, object>();
            foreach (var info in propertyInfo)
            {
                if (info.GetCustomAttributes(attributeType, true).FirstOrDefault() is not ExcelifyRecordAttribute attribute)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(attribute.FieldName))
                    propertyNames.Add(info.Name, attribute.FieldName);
                else
                {
                    propertyNames.Add(info.Name, attribute.FieldPosition);
                }
            }
            return propertyNames;
        }


        /// <summary>
        ///  Get the property name and attribute name
        /// </summary>
        /// <typeparam name="TAttribute">Attribute field</typeparam>
        /// <param name="t">Type of the object to read</param>
        /// <returns>A dictionary</returns>
        /// <exception cref="NullReferenceException"></exception>
        public static List<ExcelifyProperty> GetAttributeProperty<TAttribute>(Type t) where TAttribute : ExcelifyRecordAttribute
        {
            var propertyInfo = t.GetProperties();
            if (propertyInfo.Length == 0)
                throw new NullReferenceException("properties not found");

            var propertyNames = new List<ExcelifyProperty>();
            foreach (var info in propertyInfo)
            {
                if (info.GetCustomAttributes(true).FirstOrDefault() is not ExcelifyRecordAttribute attribute)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(attribute.FieldName))
                    propertyNames.Add(new ExcelifyProperty(info.Name, attribute.FieldName,info.PropertyType));
                else
                {
                    propertyNames.Add(new ExcelifyProperty(info.Name, attribute.FieldPosition, info.PropertyType));
                }
            }
            return propertyNames;
        }

        public static List<ExcelifyProperty> GetAttributeProperty<TAttribute, TEntity>()
            where TAttribute : ExcelifyRecordAttribute
            where TEntity : class
        {
            var propertyInfo = typeof(TEntity).GetProperties();
            if (propertyInfo.Length == 0)
                throw new NullReferenceException("properties not found");

            var propertyNames = new List<ExcelifyProperty>();
            foreach (var info in propertyInfo)
            {
                if (info.GetCustomAttributes(true).FirstOrDefault() is not ExcelifyRecordAttribute attribute)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(attribute.FieldName))
                    propertyNames.Add(new ExcelifyProperty(info.Name, attribute.FieldName, info.PropertyType));
                else
                {
                    propertyNames.Add(new ExcelifyProperty(info.Name, attribute.FieldPosition, info.PropertyType));
                }
            }
            return propertyNames;
        }

        public static List<ExcelifyProperty> GetAttributeProperty(Type attributeType, Type entityType)
        {
            var propertyInfo = entityType.GetProperties();
            if (propertyInfo.Length == 0)
                throw new NullReferenceException("properties not found");

            var propertyNames = new List<ExcelifyProperty>();
            foreach (var info in propertyInfo)
            {
                if (info.GetCustomAttributes(attributeType, true).FirstOrDefault() is not ExcelifyRecordAttribute attribute)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(attribute.FieldName))
                    propertyNames.Add(new ExcelifyProperty(info.Name, attribute.FieldName, info.PropertyType));
                else
                {
                    propertyNames.Add(new ExcelifyProperty(info.Name, attribute.FieldPosition, info.PropertyType));
                }
            }
            return propertyNames;
        }
    }
}
