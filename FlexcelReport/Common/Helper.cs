using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Report.Common
{
    public static partial class Helper
    {
        #region Constants

        public static readonly byte[] EmptyData = new byte[] { };
        public static readonly long[] EmptyIds = new long[] { };
        public static readonly object[] EMPTY_PARAMS = new object[] { };
        public static readonly string[] EMPTY_STRINGS = new string[] { };
        public static readonly char[] NEWLINES = new char[] { '\n', '\r' };
        public static readonly string[] NEWLINES_ENV = new string[] { Environment.NewLine };
        public static readonly char[] SLASHES = new char[] { '\\', '/' };
        public static readonly char[] SLASHES_FORWARD = new char[] { '/' };
        public static readonly char[] COMMAS = new char[] { ',' };
        public static readonly char[] SEMICOLONS = new char[] { ';' };
        public static readonly char[] COLONS = new char[] { ':' };
        public static readonly char[] DOTS = new char[] { '.' };
        public static readonly char[] CROSS = new char[] { '-' };
        public static readonly char[] SPACES = new char[] { ' ' };
        public static readonly char[] BLANKS = new char[] { ' ', '\t', '\n', '\r' };
        public static readonly string[] COMMA_SPLITS = new string[] { ", " };

        public static readonly System.Reflection.Assembly ASM_RUNNING = Assembly.GetExecutingAssembly();
        public static readonly AssemblyName ASM_ONAME = ASM_RUNNING.GetName();
        public static readonly string ASM_FULLNAME = ASM_ONAME.FullName;
        public static readonly System.Version APP_VERSION = ASM_ONAME.Version;

        public static readonly string APP_PATH = new Uri(ASM_RUNNING.CodeBase).LocalPath;
        public static readonly DateTime APP_DATETIME = File.GetLastWriteTime(APP_PATH);
        public static readonly string APP_NAME = Path.GetFileNameWithoutExtension(APP_PATH);
        public static readonly string APP_EXT = Path.GetExtension(APP_PATH);
        public static readonly string APP_NAMEEXT = Path.GetFileName(APP_PATH);
        public static readonly string APP_FOLDER = Path.GetDirectoryName(APP_PATH);

        public static readonly HashSet<string> EmptyHash = new HashSet<string>();

        public static readonly string TEMP_PATH = Path.GetTempPath();

        private static readonly Random GLOBAL_RANDOM = new Random();

        #endregion

        #region XML

        public static T DeserializeFile<T>(this XmlSerializer serializer, string path, out Exception exception)
            where T : class
        {
            exception = null;
            try
            {
                return DeserializeFile<T>(serializer, path);
            }
            catch (Exception ex)
            {
                exception = ex;
                return default(T);
            }
        }
        public static T DeserializeFile<T>(this XmlSerializer serializer, string path)
            where T : class
        {
            using (FileStream reader = File.OpenRead(path))
                return (T)serializer.Deserialize(reader);
        }

        public static T Deserialize<T>(this XmlSerializer serializer, byte[] data, out Exception exception)
            where T : class
        {
            exception = null;
            try
            {
                return Deserialize<T>(serializer, data);
            }
            catch (Exception ex)
            {
                exception = ex;
                return default(T);
            }
        }
        public static T Deserialize<T>(this XmlSerializer serializer, byte[] data)
            where T : class
        {
            using (MemoryStream stream = new MemoryStream(data))
                return (T)serializer.Deserialize(stream);
        }

        public static T Deserialize<T>(this XmlSerializer serializer, string data, out Exception exception)
            where T : class
        {
            exception = null;
            try
            {
                return Deserialize<T>(serializer, data);
            }
            catch (Exception ex)
            {
                exception = ex;
                return default(T);
            }
        }
        public static T Deserialize<T>(this XmlSerializer serializer, string data)
            where T : class
        {
            using (StringReader reader = new StringReader(data))
                return (T)serializer.Deserialize(reader);
        }

        public static Exception TrySerialize<T>(this XmlSerializer serializer, T t, string path)
            where T : class
        {
            try
            {
                Serialize<T>(serializer, t, path);
            }
            catch (Exception ex)
            {
                return ex;
            }
            return null;
        }
        public static void Serialize<T>(this XmlSerializer serializer, T t, string path)
            where T : class
        {
            using (FileStream writer = File.Create(path))
                serializer.Serialize(writer, t);
        }

        public static Exception TrySerialize<T>(this XmlSerializer serializer, T t)
            where T : class
        {
            try
            {
                Serialize<T>(serializer, t);
            }
            catch (Exception ex)
            {
                return ex;
            }
            return null;
        }
        public static byte[] Serialize<T>(this XmlSerializer serializer, T t)
            where T : class
        {
            using (MemoryStream stream = new MemoryStream())
            {
                serializer.Serialize(stream, t);
                return stream.GetBuffer();
            }
        }

        public static string TrySerializeText<T>(this XmlSerializer serializer, T t, out Exception exception)
            where T : class
        {
            exception = null;
            try
            {
                return SerializeText<T>(serializer, t);
            }
            catch (Exception ex)
            {
                exception = ex;
                return null;
            }
        }
        public static string SerializeText<T>(this XmlSerializer serializer, T t)
            where T : class
        {
            using (StringWriter writer = new StringWriter())
            {
                serializer.Serialize(writer, t);
                return writer.ToString();
            }
        }

        #endregion

        #region JSON

        /// <summary>
        /// Create a User object and serialize it to a JSON stream.
        /// </summary>
        public static string JsonDataSerialize(object value, Type type)
        {
            if (value == null)
                return null;

            //Create a stream to serialize the object to.
            using (var ms = new MemoryStream())
            {
                var ser = new DataContractJsonSerializer(type ?? value.GetType());
                ser.WriteObject(ms, value);
                byte[] json = ms.ToArray();
                return Encoding.UTF8.GetString(json, 0, json.Length);
            }
        }

        public static string JsonDataSerialize<T>(T t)
        {
            return JsonDataSerialize(t, typeof(T));
        }

        /// <summary>
        /// Deserialize a JSON stream to a User object.
        /// </summary>
        public static object JsonDataDeserialize(string json, Type type)
        {
            if (json == null || String.IsNullOrWhiteSpace(json))
                return null;
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                return new DataContractJsonSerializer(type).ReadObject(ms);
        }

        public static T JsonDataDeserialize<T>(string json)
        {
            return (T)JsonDataDeserialize(json, typeof(T));
        }

        #endregion

        #region ValueTypes

        //public enum TypeCode[Empty = 0 => String = 18]
        public readonly static object[] DefaultValues = new object[]
        {
            null, // Empty = 0
            default(Object), // Object = 1
            default(DBNull), // DBNull = 2
            default(Boolean), // Boolean = 3
            default(Char), // Char = 4
            default(SByte), // SByte = 5
            default(Byte), // Byte = 6
            default(Int16), // Int16 = 7
            default(UInt16), // UInt16 = 8
            default(Int32), // Int32 = 9
            default(UInt32), // UInt32 = 10
            default(Int64), // Int64 = 11
            default(UInt64), // UInt64 = 12
            default(Single), // Single = 13
            default(Double), // Double = 14
            default(Decimal), // Decimal = 15
            default(DateTime), // DateTime = 16
            null, // Missing = 17
            default(String), // String = 18
        };

        public readonly static object[] MinValues = new object[]
        {
            null, // Empty = 0
            null, // Object = 1
            null, // DBNull = 2
            false, // Boolean = 3
            Char.MinValue, // Char = 4
            SByte.MinValue, // SByte = 5
            Byte.MinValue, // Byte = 6
            Int16.MinValue, // Int16 = 7
            UInt16.MinValue, // UInt16 = 8
            Int32.MinValue, // Int32 = 9
            UInt32.MinValue, // UInt32 = 10
            Int64.MinValue, // Int64 = 11
            UInt64.MinValue, // UInt64 = 12
            Single.MinValue, // Single = 13
            Double.MinValue, // Double = 14
            Decimal.MinValue, // Decimal = 15
            DateTime.MinValue, // DateTime = 16
            null, // Missing = 17
            null, // String = 18
        };

        public readonly static object[] MaxValues = new object[]
        {
            null, // Empty = 0
            null, // Object = 1
            null, // DBNull = 2
            true, // Boolean = 3
            Char.MaxValue, // Char = 4
            SByte.MaxValue, // SByte = 5
            Byte.MaxValue, // Byte = 6
            Int16.MaxValue, // Int16 = 7
            UInt16.MaxValue, // UInt16 = 8
            Int32.MaxValue, // Int32 = 9
            UInt32.MaxValue, // UInt32 = 10
            Int64.MaxValue, // Int64 = 11
            UInt64.MaxValue, // UInt64 = 12
            Single.MaxValue, // Single = 13
            Double.MaxValue, // Double = 14
            Decimal.MaxValue, // Decimal = 15
            DateTime.MaxValue, // DateTime = 16
            null, // Missing = 17
            null, // String = 18
        };

        public readonly static Dictionary<Type, string> PrimitiveTypes = new Dictionary<Type, string>()
        {
            { typeof(System.Boolean), "bool" },
            { typeof(System.Byte), "byte" },
            { typeof(System.SByte), "sbyte" },
            { typeof(System.Char), "char" },
            { typeof(System.Decimal), "decimal" },
            { typeof(System.Double), "double" },
            { typeof(System.Single), "float" },
            { typeof(System.Int32), "int" },
            { typeof(System.UInt32), "uint" },
            { typeof(System.Int64), "long" },
            { typeof(System.UInt64), "ulong" },
            { typeof(System.Object), "object" },
            { typeof(System.Int16), "short" },
            { typeof(System.UInt16), "ushort" },
            { typeof(System.String), "string" },
        };

        public readonly static Type[] SystemTypes = new Type[]
        {
            null,
            typeof(System.Object),
            typeof(System.DBNull),
            typeof(System.Boolean),
            typeof(System.Char),
            typeof(System.SByte),
            typeof(System.Byte),
            typeof(System.Int16),
            typeof(System.UInt16),
            typeof(System.Int32),
            typeof(System.UInt32),
            typeof(System.Int64),
            typeof(System.UInt64),
            typeof(System.Single),
            typeof(System.Double),
            typeof(System.Decimal),
            typeof(System.DateTime),
            null,
            typeof(System.String),
            typeof(System.Guid),
            //typeof(Core.Month),
            //typeof(Core.Money),
            //typeof(Core.LText),
        };

        public readonly static System.ComponentModel.TypeConverter[] SystemConverters = new System.ComponentModel.TypeConverter[]
        {
            null,
            null,
            null,
            new System.ComponentModel.BooleanConverter(),
            new System.ComponentModel.CharConverter(),
            new System.ComponentModel.SByteConverter(),
            new System.ComponentModel.ByteConverter(),
            new System.ComponentModel.Int16Converter(),
            new System.ComponentModel.UInt16Converter(),
            new System.ComponentModel.Int32Converter(),
            new System.ComponentModel.UInt32Converter(),
            new System.ComponentModel.Int64Converter(),
            new System.ComponentModel.UInt64Converter(),
            new System.ComponentModel.SingleConverter(),
            new System.ComponentModel.DoubleConverter(),
            new System.ComponentModel.DecimalConverter(),
            new System.ComponentModel.DateTimeConverter(),
            null,
            new System.ComponentModel.StringConverter(),
            new System.ComponentModel.GuidConverter(),
            null, // need to init by System Framework
        };

        #endregion ValueTypes

        #region Random

        public static IEnumerable<int> RandomArrayIndexes(int arrayLength, int items)
        {
            return items > (arrayLength / 2) ?
                RandomArrayIndexesReserved(arrayLength, items) :
                RandomArrayIndexesSequenced(arrayLength, items);
        }

        private static IEnumerable<int> RandomArrayIndexesReserved(int arrayLength, int items)
        {
            var arrayIndexes = new int[arrayLength];
            for (var i = 0; i < arrayLength; i++)
                arrayIndexes[i] = i;
            var listIndexes = new List<int>(arrayIndexes);
            for (var i = 0; i < items; i++)
            {
                var index = GLOBAL_RANDOM.Next(listIndexes.Count);
                var value = listIndexes[index];
                listIndexes.RemoveAt(index);
                yield return value;
            }
        }

        private static IEnumerable<int> RandomArrayIndexesSequenced(int arrayLength, int items)
        {
            var hash = new HashSet<int>();
            for (var i = 0; i < items; i++)
            {
                int index;
                for (index = GLOBAL_RANDOM.Next(arrayLength);
                    !hash.Add(index);
                    index = GLOBAL_RANDOM.Next(arrayLength)) ;
                yield return index;
            }
        }

        public static string RandomPassword(int length)
        {
            // Initiate objects & vars
            Random random = new Random();
            StringBuilder randomString = new StringBuilder(length);
            int randNumber;

            // Loop length times to generate a random number or character
            for (int i = 0; i < length; i++)
            {
                if (random.Next(1, 3) == 1)
                    randNumber = random.Next(97, 123); //char {a-z}
                else
                    randNumber = random.Next(48, 58); //int {0-9}
                randomString.Append((char)randNumber);
            }
            return randomString.ToString();
        }

        #endregion

        #region Dynamic

        /// <summary>
        /// Gán giá trị động cho trường readonly của lớp
        /// </summary>
        public static T DynamicSet<T>(Type type, string field, T value)
        {
            type.GetField(field, DynamicDelegates.ClassBindingFlags).SetValue(null, value);
            return value;
        }

        /// <summary>
        /// Gán giá trị động cho trường readonly của object
        /// </summary>
        public static TVal DynamicSet<TObj, TVal>(TObj instance, string field, TVal value)
        {
            typeof(TObj).GetField(field, DynamicDelegates.ObjectBindingFlags).SetValue(instance, value);
            return value;
        }

        /// <summary>
        /// Gán giá trị động cho trường readonly của object
        /// </summary>
        public static T DynamicSet<T>(Type type, string field, object instance, T value)
        {
            type.GetField(field, DynamicDelegates.ObjectBindingFlags).SetValue(instance, value);
            return value;
        }

        /// <summary>
        /// Gán giá trị động cho trường readonly của lớp/object
        /// </summary>
        public static T DynamicSet<T>(Expression<Func<T>> selector, T value)
        {
#if DEBUG
            if (selector.Body.NodeType != ExpressionType.MemberAccess)
                throw new ArgumentException(String.Format("[{1}] hàm chọn trường bị sai: {0}", selector.Body.NodeType, selector), "selector");
#endif
            var left = selector.Body;
            var inner = (left as MemberExpression).Expression;
            while (inner != null && inner is MemberExpression)
            {
                left = inner;
                inner = (inner as MemberExpression).Expression;
            }
            ((selector.Body as MemberExpression).Member as FieldInfo).SetValue(
                //inner == null || !(inner is ConstantExpression) ? null : (inner as ConstantExpression).Value.GetType().GetField((left as MemberExpression).Member.Name, DynamicDelegates.ObjectBindingFlags).GetValue((inner as ConstantExpression).Value),
                inner == null || !(inner is ConstantExpression) ?
                    // Class or Object?
                    null :
                    (inner as ConstantExpression).Value.GetType().IsDefined(typeof(System.Runtime.CompilerServices.CompilerGeneratedAttribute), false) ?
                        // is direct value or in direct through a anonymous generated type?
                        (inner as ConstantExpression).Value.GetType().GetField((left as MemberExpression).Member.Name, DynamicDelegates.ObjectBindingFlags).GetValue((inner as ConstantExpression).Value) :
                        (inner as ConstantExpression).Value
                    ,
                value);

            return value;
        }

        public static string DynamicFieldName<TObject, TField>(Expression<Func<TObject, TField>> selector)
        {
#if DEBUG
            if (selector.Body.NodeType != ExpressionType.MemberAccess)
                throw new ArgumentException(String.Format("[{1}] hàm chọn trường bị sai: {0}", selector.Body.NodeType, selector), "selector");
#endif
            return ((MemberExpression)selector.Body).Member.Name;
        }

        public static Array DynamicToArray(Type elementType, IEnumerable data)
        {
            var objects = data.Cast<object>().ToArray();
            var length = objects.Length;
            var elements = Array.CreateInstance(elementType, length);
            for (var index = 0; index < length; index++)
                elements.SetValue(objects[index], index);
            return elements;
        }

        //public static Array DynamicToArray(this ObjectResult result)
        //{
        //    return Helper.DynamicToArray(result.ElementType, result);
        //}

        public static Array DynamicToArray(this IQueryable query)
        {
            return Helper.DynamicToArray(query.ElementType, query);
        }

        public static object DynamicToListSource(object value)
        {
            object result;
            if (value == null)
                result = null;
            else if (value is System.Collections.IEnumerable)
                result = value;
            else
                result = new object[] { new { Value = value } };

            return result;
        }

        public static object DynamicSingle(System.Collections.IEnumerable value)
        {
            if (value == null)
                return null;
            var enumerator = value.GetEnumerator();
            return enumerator.MoveNext() ? enumerator.Current : null;
        }

        public static object DynamicSingle(System.Dynamic.ExpandoObject value)
        {
            return ((IEnumerable<KeyValuePair<string, object>>)value).Select(kv => kv.Value).SingleOrDefault();
        }

        public static Func<TObject, TProperty> DynamicPropertyGetter<TObject, TProperty>(string propertyName)
        {
            var paramExpression = Expression.Parameter(typeof(TObject), "value");
            var propertyGetterExpression = Expression.Property(paramExpression, propertyName);
            return Expression.Lambda<Func<TObject, TProperty>>(propertyGetterExpression, paramExpression).Compile();
        }

        public static Action<TObject, TProperty> DynamicPropertySetter<TObject, TProperty>(string propertyName)
        {
            var paramExpression = Expression.Parameter(typeof(TObject));
            var paramExpression2 = Expression.Parameter(typeof(TProperty), propertyName);
            var propertyGetterExpression = Expression.Property(paramExpression, propertyName);

            return Expression.Lambda<Action<TObject, TProperty>>(
                Expression.Assign(propertyGetterExpression, paramExpression2),
                paramExpression, paramExpression2).Compile();
        }

        public static Func<TObject, TField> DynamicFieldGetter<TObject, TField>(string fieldName)
        {
            var paramExpression = Expression.Parameter(typeof(TObject), "value");
            var fieldGetterExpression = Expression.Field(paramExpression, fieldName);
            return Expression.Lambda<Func<TObject, TField>>(fieldGetterExpression, paramExpression).Compile();
        }

        public static Action<TObject, TField> DynamicFieldSetter<TObject, TField>(string fieldName)
        {
            var paramExpression = Expression.Parameter(typeof(TObject));
            var paramExpression2 = Expression.Parameter(typeof(TField), fieldName);
            var propertyGetterExpression = Expression.Field(paramExpression, fieldName);

            return Expression.Lambda<Action<TObject, TField>>(
                Expression.Assign(propertyGetterExpression, paramExpression2),
                paramExpression, paramExpression2).Compile();
        }

        #endregion

        #region ToXXX

        //
        // Summary:
        //     Creates an array from a System.Collections.Generic.IEnumerable<T> or return Null if source is empty.
        //
        // Parameters:
        //   source:
        //     An System.Collections.Generic.IEnumerable<T> to create an array from.
        //
        // Type parameters:
        //   TSource:
        //     The type of the elements of source.
        //
        // Returns:
        //     An array that contains the elements from the input sequence.
        //
        // Exceptions:
        //   System.ArgumentNullException:
        //     source is null.
        public static TSource[] ToArrayOrNull<TSource>(this IEnumerable<TSource> source)
        {
            return !source.Any() ? null : source.ToArray();
        }

        public static Dictionary<TKey, TElement> ToDictionaryOrNull<TSource, TKey, TElement>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector, Func<TSource, TElement> elementSelector)
        {
            return !source.Any() ? null : source.ToDictionary(keySelector, elementSelector);
        }

        public static Dictionary<TKey, TSource> ToDictionaryWithCheck<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
#if DEBUG
            try
            {
                return source.ToDictionary(keySelector);
            }
            catch (ArgumentException ex) // Duplicate keys for two elements
            {
                throw new ArgumentException(
                    String.Format("\n\n{0} có các mã sau bị trùng: \n    {1}\n\n", typeof(TSource), source.GroupBy(keySelector).Where(i => i.Count() > 1).Select((i, n) => String.Format("{0}. {1} [{2}]", n + 1, i.Key, i.Select(o => o == null ? "<null>" : o.ToString()).Join(","))).Join(";\n    ")),
                    ex);
            }
#else
            return source.ToDictionary(keySelector);
#endif
        }

        public static Dictionary<TKey, TElement> ToDictionaryWithCheck<TSource, TKey, TElement>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector, Func<TSource, TElement> elementSelector)
        {
#if DEBUG
            try
            {
                return source.ToDictionary(keySelector, elementSelector);
            }
            catch (ArgumentException ex) // Duplicate keys for two elements
            {
                throw new ArgumentException(
                    String.Format("\n\n{0} có các mã sau bị trùng: \n    {1}\n\n", typeof(TSource), source.GroupBy(keySelector).Where(i => i.Count() > 1).Select((i, n) => String.Format("{0}. {1} [{2}]", n + 1, i.Key, i.Select(o => o == null ? "<null>" : o.ToString()).Join(","))).Join(";\n    ")),
                    ex);
            }
#else
            return source.ToDictionary(keySelector, elementSelector);
#endif
        }

        public static Dictionary<TBase, string> ToDictionary<TEnum, TBase>(Func<TEnum, TBase> keySelector, Func<TEnum, string> elementSelector = null)
        {
            return Enum.GetValues(typeof(TEnum)).Cast<TEnum>().ToDictionary(keySelector, elementSelector == null ? e => e.ToString() : elementSelector);
        }

        public static Dictionary<TBase, string> ToDictionary<TEnum, TBase>(Func<TEnum, string> elementSelector = null)
        {
            return Enum.GetValues(typeof(TEnum)).Cast<TEnum>().ToDictionary(e => (TBase)Convert.ChangeType(e, typeof(TBase)), elementSelector == null ? e => e.ToString() : elementSelector);
        }

        public static Int32 ToInt32(this object strNumber)
        {
            return strNumber.ToInt32(0);
        }

        public static Int32 ToInt32(this object strNumber, Int32 _default)
        {
            Int32 result;
            return strNumber != null && Int32.TryParse(strNumber is String ? (String)strNumber : strNumber.ToString(), out result) ? result : _default;
        }

        public static Int32? ToNullableInt32(string strNumber)
        {
            Int32 result;
            return strNumber != null && Int32.TryParse((string)strNumber, out result) ? (Int32?)result : null;
        }

        public static Int32? ToInt32Ptr(this object strNumber)
        {
            Int32 result;
            return strNumber != null && Int32.TryParse(strNumber is String ? (String)strNumber : strNumber.ToString(), out result) ? (Int32?)result : null;
        }

        public static Int64 ToInt64(this object strNumber)
        {
            return strNumber.ToInt64(0);
        }

        public static Int64 ToInt64(this object strNumber, Int64 _default)
        {
            Int64 result;
            return strNumber != null && Int64.TryParse(strNumber is String ? (String)strNumber : strNumber.ToString(), out result) ? result : _default;
        }

        public static Int64? ToInt64Ptr(this object strNumber)
        {
            Int64 result;
            return strNumber != null && Int64.TryParse(strNumber is String ? (String)strNumber : strNumber.ToString(), out result) ? (Int64?)result : null;
        }

        public static E ToEnum<E>(string value, E _default)
        {
            if (String.IsNullOrEmpty(value))
                return _default;
            try
            {
                return (E)Enum.Parse(typeof(E), value, true);
            }
            catch
            {
                return _default;
            }
        }

        public static Int64 ToInt64(Version version)
        {
            return ((long)version.Major) * 100L * 100000L * 100L + ((long)version.Minor) * 100000L * 100L + ((long)version.Build) * 100L + (long)version.Revision;
        }

        public static Int64 ToInt64(Version version1, Version version2)
        {
            return
                (((long)version1.Major) * 10L * 100000L * 10L + ((long)version1.Minor) * 100000L * 10L + ((long)version1.Build) * 10L + (long)version1.Revision) * 100L * 10L * 100000L * 10L +
                (((long)version2.Major) * 10L * 100000L * 10L + ((long)version2.Minor) * 100000L * 10L + ((long)version2.Build) * 10L + (long)version2.Revision);
        }

        #endregion

        #region Join strings

        public static string Join(this IEnumerable<string> values, string separator)
        {
            StringBuilder sb = new StringBuilder(0x100);
            foreach (string value in values)
                sb.Append(value).Append(separator);
            return sb.Length <= 0 ? String.Empty : sb.ToString(0, sb.Length - separator.Length);
        }

        public static string JoinOrNull(this IEnumerable<string> values, string separator)
        {
            if (!values.Any())
                return null;
            StringBuilder sb = new StringBuilder(0x100);
            foreach (string value in values)
                sb.Append(value).Append(separator);
            return sb.ToString(0, sb.Length - separator.Length);
        }

        public static string Join(this string[] values, string separator)
        {
            return String.Join(separator, values);
        }

        public static string Join<T>(this T[] values)
        {
            return values == null ? String.Empty : values.Length == 1 ? values[0].ToString() : values.Join(COMMAS[0].ToString(), 0, values.Length).ToString();
        }
        public static string Join<T>(this T[] values, string separator)
        {
            return values == null ? String.Empty : values.Length == 1 ? values[0].ToString() : values.Join(separator, 0, values.Length).ToString();
        }
        public static string Join<T>(this T[] values, int from, int length)
        {
            return values.Join(COMMAS[0].ToString(), from, length).ToString();
        }
        public static StringBuilder Join<T>(this T[] values, string separator, int from, int length)
        {
            return values.Join(separator, "{0}", from, length);
        }
        public static StringBuilder Join<T>(this T[] values, string separator, String valueFormat, int from, int length)
        {
            if (values == null || values.Length <= 0)
                return new StringBuilder();
            StringBuilder sb = new StringBuilder();
            for (int i = from; i < from + length; i++)
                sb.AppendFormat(valueFormat + "{1}", values[i], i < (values.Length - 1) ? separator : String.Empty);
            return sb;
        }

        public static StringBuilder JoinLite<T>(this IEnumerable<T> values)
        {
            return values.JoinLite(Helper.COMMAS[0], 0, 0);
        }
        public static StringBuilder JoinLite<T, TSep>(this IEnumerable<T> values, TSep separator)
        {
            return values.JoinLite(separator, 0, 0);
        }
        public static StringBuilder JoinLite<T, TSep>(this IEnumerable<T> values, TSep separator, int from, int length)
        {
            var sepstring = separator as string ?? separator.ToString();
            if (from > 0)
                values = values.Skip(from);
            if (length > 0)
                values = values.Take(length);

            Func<StringBuilder, T, StringBuilder> func;
            if (typeof(T) == typeof(string))
                func = (sb, v) => v == null ? sb : sb.Append(v).Append(sepstring);
            else if (typeof(T) == typeof(Guid))
                func = (sb, v) => v == null ? sb : sb.Append(((Guid)(object)v).ToString("D")).Append(sepstring);
            else
                func = (sb, v) => v == null ? sb : sb.Append(v.ToString()).Append(sepstring);

            return values.Aggregate(new StringBuilder(), func, sb => sb.Length <= 0 ? sb : sb.Remove(sb.Length - sepstring.Length, sepstring.Length));
        }

        public static string JoinToString<T>(this IEnumerable<T> values)
        {
            if (values == null)
                return "*";
            else if (values.Count() <= 0)
                return "[]";
            else
                return values.JoinLite(Helper.COMMAS[0], 0, 0).Insert(0, '[').Append(']').ToString();
        }

        public static IEnumerable<TResut> JoinParseLite<T, TResut>(this string joint)
        {
            if (joint == null || joint.Length <= 0 || String.IsNullOrWhiteSpace(joint) ||
                (joint.Length == 1 && joint[0] == '*'))
                return null;
            joint = joint.Substring(1, joint.Length - 2);
            if (joint == null || joint.Length <= 0 || String.IsNullOrWhiteSpace(joint))
                return YieldEmpty<TResut>();
            var strings = joint.Split(COMMAS, StringSplitOptions.RemoveEmptyEntries);
            if (strings == null || strings.Length <= 0)
                return YieldEmpty<TResut>();
            return typeof(T) == typeof(string) ?
                strings.Cast<TResut>() :
                typeof(T) == typeof(Guid) ?
                    strings.Select(s => Guid.ParseExact(s, "D")).Cast<TResut>() :
                    strings.Select(s => SafeConvert<T>(s)).Cast<TResut>();
        }

        public static T[] JoinParse<T>(this string joint)
        {
            var result = JoinParseLite<T, T>(joint);
            return result == null ? null : result.ToArray();
        }

        public static TResut[] JoinParse<T, TResut>(this string joint)
        {
            var result = JoinParseLite<T, TResut>(joint);
            return result == null ? null : result.ToArray();
        }

        #endregion

        #region SafeParse

        public static T SafeParse<T>(this string input)
        {
            if (input == null || input.Length <= 0)
                return default(T);
            if (typeof(T) == typeof(string))
                return (T)(object)input;
            if (typeof(T) == typeof(Guid))
                return (T)(object)Guid.ParseExact(input, "D");
            return SafeConvert<T>(input);
        }

        public static T SafeConvert<T>(this string input)
        {
            var converter = System.ComponentModel.TypeDescriptor.GetConverter(typeof(T));
            if (converter != null)
                try { return (T)converter.ConvertFromString(input); }
                catch { }
            return default(T);
        }

        public static object SafeConvertFromString(string input, Type type)
        {
            if (type == typeof(string))
                return input;
            if (type.IsGenericType)
                if (type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    type = type.GetGenericArguments()[0];
                else
                    return null;
            var typecode = (int)Type.GetTypeCode(type);
            return (typecode < 0 || typecode >= SystemConverters.Length || SystemConverters[typecode] == null) ? null :
                SystemConverters[typecode].ConvertFromString(input);
        }

        #endregion

        #region Exceptions

#if false
        private static readonly DynamicObjectAction<Exception> Exception_PrepForRemoting = DynamicDelegates.CreateObjectMethod<DynamicObjectAction<Exception>>("PrepForRemoting");
#endif
        public static Exception PrepForRemoting(this Exception wrapper)
        {
#if true
            typeof(Exception).InvokeMember(
                "PrepForRemoting",
                BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.InvokeMethod,
                (Binder)null, wrapper, Helper.EMPTY_PARAMS);
            return wrapper;
#else
#if true
            Exception_PrepForRemoting(wrapper);
            return wrapper;
#else
            var ctx = new StreamingContext(StreamingContextStates.CrossAppDomain);
            var mgr = new ObjectManager(null, ctx);
            var si = new SerializationInfo(wrapper.GetType(), new FormatterConverter());

            wrapper.GetObjectData(si, ctx);
            mgr.RegisterObject(wrapper, 1, si); // prepare for SetObjectData
            mgr.DoFixups(); // ObjectManager calls SetObjectData

            // voila, e is unmodified save for _remoteStackTraceString
            return wrapper.InnerException;
#endif
#endif
        }

        public static Exception FindInnerException(this Exception exception)
        {
            return exception.InnerException == null ? exception : exception.InnerException.FindInnerException();
        }

        public static Exception OtherInnerException(this Exception exception)
        {
            var inner = exception.InnerException;
            return inner == null ? null :
                inner.GetType() == exception.GetType() ?
                    inner.OtherInnerException() :
                    inner;
        }

        #endregion

        #region Resources

        public static byte[] GetResourceData(Assembly assembly, string resourceName)
        {
            if (assembly.GetManifestResourceInfo(resourceName) == null)
                return null;
            return StreamHelper.Read(assembly.GetManifestResourceStream(resourceName));
        }

        public static string GetResourceString(Assembly assembly, string resourceName)
        {
            if (assembly.GetManifestResourceInfo(resourceName) == null)
                return null;
            return Encoding.UTF8.GetString(StreamHelper.Read(assembly.GetManifestResourceStream(resourceName)));
        }

        #endregion

        #region String Utils

        public static string FormatWith(this string format, IFormatProvider provider, params object[] args)
        {
            return string.Format(provider, format, args);
        }

        /// <summary>
        /// Determines whether the string is all white space. Empty string will return false.
        /// </summary>
        /// <param name="s">The string to test whether it is all white space.</param>
        /// <returns>
        /// 	<c>true</c> if the string is all white space; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsWhiteSpace(string s)
        {
            if (s == null)
                throw new ArgumentNullException("s");

            if (s.Length == 0)
                return false;

            for (int i = 0; i < s.Length; i++)
            {
                if (!char.IsWhiteSpace(s[i]))
                    return false;
            }

            return true;
        }

#if !LAUNCHER
        public static string ToString(this System.Data.Common.DbDataRecord dataValue, int fieldIndex)
        {
            if (dataValue[fieldIndex] is String)
                return (String)dataValue[fieldIndex];
            else if (dataValue[fieldIndex].GetType().IsPrimitive)
                return dataValue[fieldIndex].ToString();
            else if (dataValue[fieldIndex] is System.Data.Common.DbDataRecord)
                return ((System.Data.Common.DbDataRecord)dataValue[fieldIndex]).ToString(fieldIndex);
            return null;
        }
#endif

        #endregion

        #region Yield

        public static IEnumerable<T> YieldEmpty<T>()
        {
            yield break;
        }

        public static IEnumerable<T> YieldSingle<T>(T t)
        {
            yield return t;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2)
        {
            yield return t1;
            yield return t2;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3)
        {
            yield return t1;
            yield return t2;
            yield return t3;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3, T t4)
        {
            yield return t1;
            yield return t2;
            yield return t3;
            yield return t4;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3, T t4, T t5)
        {
            yield return t1;
            yield return t2;
            yield return t3;
            yield return t4;
            yield return t5;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3, T t4, T t5, T t6)
        {
            yield return t1;
            yield return t2;
            yield return t3;
            yield return t4;
            yield return t5;
            yield return t6;
        }


        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3, T t4, T t5, T t6, T t7)
        {
            yield return t1;
            yield return t2;
            yield return t3;
            yield return t4;
            yield return t5;
            yield return t6;
            yield return t7;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3, T t4, T t5, T t6, T t7, T t8)
        {
            yield return t1;
            yield return t2;
            yield return t3;
            yield return t4;
            yield return t5;
            yield return t6;
            yield return t7;
            yield return t8;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3, T t4, T t5, T t6, T t7, T t8, T t9)
        {
            yield return t1;
            yield return t2;
            yield return t3;
            yield return t4;
            yield return t5;
            yield return t6;
            yield return t7;
            yield return t8;
            yield return t9;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3, T t4, T t5, T t6, T t7, T t8, T t9, T t10)
        {
            yield return t1;
            yield return t2;
            yield return t3;
            yield return t4;
            yield return t5;
            yield return t6;
            yield return t7;
            yield return t8;
            yield return t9;
            yield return t10;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3, T t4, T t5, T t6, T t7, T t8, T t9, T t10, T t11)
        {
            yield return t1;
            yield return t2;
            yield return t3;
            yield return t4;
            yield return t5;
            yield return t6;
            yield return t7;
            yield return t8;
            yield return t9;
            yield return t10;
            yield return t11;
        }

        public static IEnumerable<T> YieldList<T>(T t1, T t2, T t3, T t4, T t5, T t6, T t7, T t8, T t9, T t10, T t11, T t12)
        {
            yield return t1;
            yield return t2;
            yield return t3;
            yield return t4;
            yield return t5;
            yield return t6;
            yield return t7;
            yield return t8;
            yield return t9;
            yield return t10;
            yield return t11;
            yield return t12;
        }

        #endregion

        #region Reflections

        public static System.Drawing.Image CopyImage(this System.Drawing.Image image)
        {
            return image == null ? null : new System.Drawing.Bitmap(image);
        }

        public static System.IO.Stream ToStream(this System.Drawing.Image image, System.Drawing.Imaging.ImageFormat imageFormat = null)
        {
            var stream = new System.IO.MemoryStream();
            image.Save(stream, imageFormat ?? System.Drawing.Imaging.ImageFormat.Bmp);
            stream.Seek(0, System.IO.SeekOrigin.Begin);
            return stream;
        }

        public static byte[] ToData(this System.Drawing.Image image, System.Drawing.Imaging.ImageFormat imageFormat = null)
        {
            using (var stream = new System.IO.MemoryStream())
            {
                image.Save(stream, imageFormat ?? System.Drawing.Imaging.ImageFormat.Bmp);
                return stream.ToArray();
            }
        }

        public static bool IsEquals(byte[] source, byte[] target)
        {
            if (source == null)
                return target == null;
            if (target == null)
                return source == null;
            if (source.Length != target.Length)
                return false;
            if (source.Length <= 0)
                return true;
            for (int i = 0, l = source.Length; i < l; i++)
                if (source[i] != target[i])
                    return false;
            return true;
        }

        public static IEnumerable<Type> EnumerateBaseTypes(Type type, Type mostBaseType)
        {
            for (var currentType = type; currentType.BaseType != null && currentType != mostBaseType; currentType = currentType.BaseType)
                yield return currentType;
        }

        public static IEnumerable<Type> EnumerateBaseTypes(Type type, HashSet<Type> mostBaseTypes)
        {
            for (var currentType = type; currentType.BaseType != null && !mostBaseTypes.Contains(currentType); currentType = currentType.BaseType)
                yield return currentType;
        }

        public static Type LoadType(string typeName, string assemblyName)
        {
            var assembly = Assembly.LoadFrom(assemblyName);
            return assembly == null ? null : assembly.GetType(typeName, false);
        }

        public static Object CreateInstance(String strAssembly, String strClass, Object[] args)
        {
            if (strClass.Length <= 0)
                return null;

            Assembly asm = strAssembly.Length <= 0 ? Assembly.GetExecutingAssembly() : Assembly.LoadFrom(strAssembly);
            Type type = asm.GetType(strClass, true, true);
            return Activator.CreateInstance(type, args);
        }

        public static IEnumerable<Type> FindConcreateTypes(Type type, IEnumerable<Assembly> assemblies)
        {
            foreach (var subType in assemblies.SelectMany(asm => asm.GetTypes().Where(t => t.BaseType == type)))
                if (!subType.IsAbstract)
                    yield return subType;
                else
                    foreach (var concreateType in Helper.FindConcreateTypes(subType, assemblies))
                        yield return concreateType;
        }

        public static bool IsNullable(this Type type)
        {
            return type.GUID == GUID_Nullable;
        }
        static readonly Guid GUID_Nullable = new Guid("(9a9177c7-cf5f-31ab-8495-96f58ac5df3a)");

        private static SynchronizedDictionary<string, string> AssemblyNamePathMap =
            new SynchronizedDictionary<string, string>(30, StringComparer.InvariantCultureIgnoreCase);

        /// <summary>
        /// Get the short type definition for using in code generation
        /// </summary>
        public static string GetTypeDefinition(Type type)
        {
            return Helper.GetTypeDefinition(type, false, false, null, null);
        }

        /// <summary>
        /// Get the short type definition for using in code generation
        /// </summary>
        public static string GetTypeDefinition(Type type, HashSet<string> imports, HashSet<string> assemblies)
        {
            return Helper.GetTypeDefinition(type, false, false, imports, assemblies);
        }

        public static string GetTypeDefinition(Type type, bool fullname, HashSet<string> imports, HashSet<string> assemblies)
        {
            return Helper.GetTypeDefinition(type, fullname, false, imports, assemblies);
        }

        /// <summary>
        /// Get the full/short type definition for using in code generation
        /// </summary>
        public static string GetTypeDefinition(Type type, bool fullname, bool nullable, HashSet<string> imports, HashSet<string> assemblies)
        {
            if (type == null)
                return "void";
            else if (Helper.PrimitiveTypes.ContainsKey(type))
                return !nullable || type.IsClass ? Helper.PrimitiveTypes[type] : (Helper.PrimitiveTypes[type] + "?");
            else
            {
                // Assembly
                string assemblyRef = Helper.GetAssemblyReference(type);
                if (assemblies != null && !assemblies.Contains(assemblyRef))
                    assemblies.Add(assemblyRef);

                // Import
                string namespaceName = type.Namespace;
                if (imports != null && !imports.Contains(namespaceName))
                    imports.Add(namespaceName);

                // Return type name
                string typename = fullname ? type.FullName : type.Name;
                if (type.IsArray)
                {
                    return Helper.GetTypeDefinition(type.GetElementType(), imports, assemblies) + "[]";
                }
                else if (type.IsGenericType)
                {
                    if (type.GetGenericTypeDefinition() == typeof(Nullable<>))
                        return String.Format("{0}?", GetTypeDefinition(type.GetGenericArguments().Single(), imports, assemblies));
                    else
                        return String.Format("{0}<{1}>",
                                typename.Substring(0, typename.Length - 2),
                                type.GetGenericArguments().Select(t => GetTypeDefinition(t, imports, assemblies)).Join(", ")
                                );
                }
                return nullable ? (typename + "?") : typename;
            }
        }

        public static string GetAssemblyReference(Type assemblyClass)
        {
            return Helper.GetAssemblyReference(assemblyClass.Assembly.GetName().FullName);
        }

        public static string GetAssemblyReference(string shortAssemblyName)
        {
            if (Helper.AssemblyNamePathMap.ContainsKey(shortAssemblyName))
                return Helper.AssemblyNamePathMap[shortAssemblyName];

            // Try 1st to determine this passed string is in form of AssemblyName
            try
            {
                AssemblyName assemblyName = new AssemblyName(shortAssemblyName);
                Assembly asmByName = Assembly.Load(assemblyName);
                if (asmByName != null)
                    return Helper.AssemblyNamePathMap[shortAssemblyName] = asmByName.Location;
            }
            catch
            {
            }

            // First, get the path for this executing assembly.
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // if the file exists in this Path - prepend the path
            string fullReference = Path.Combine(path, shortAssemblyName);
            if (File.Exists(fullReference))
                return Helper.AssemblyNamePathMap[shortAssemblyName] = fullReference;
            else
            {
                // Strip off any trailing ".dll" if present.
                fullReference = Path.GetFileNameWithoutExtension(shortAssemblyName);

                // See if the required assembly is already present in our current AppDomain
                foreach (Assembly currAssembly in AppDomain.CurrentDomain.GetAssemblies())
                {
                    if (string.Compare(currAssembly.GetName().Name, fullReference, true) == 0)
                    {
                        // Found it, return the location as the full reference.
                        return Helper.AssemblyNamePathMap[shortAssemblyName] = currAssembly.Location;
                    }
                }

                // The assembly isn't present in our current application, so attempt to
                // load it from the GAC, using the partial name.
                try
                {
                    Assembly tempAssembly = Assembly.Load(fullReference);
                    return Helper.AssemblyNamePathMap[shortAssemblyName] = tempAssembly.Location;
                }
                catch
                {
                    // If we cannot load or otherwise access the assembly from the GAC then just
                    // return the relative reference and hope for the best.
                    return Helper.AssemblyNamePathMap[shortAssemblyName] = shortAssemblyName;
                }
            }
        }

        public static bool IsEquals<T>(T[] arr1, T[] arr2)
        {
            for (int i = 0; i < arr1.Length; i++)
                if (!arr1[i].Equals(arr2[i]))
                    return false;
            return true;
        }

        public static string PrintMethodParameters(this MethodBase method)
        {
            return method.GetParameters().Select(p => p.Name[0].ToString().ToUpperInvariant() + p.Name.Substring(1) + ":" + p.ParameterType.Name).Join(", ");
        }

        #endregion

        #region WIN32 Securities

        public static System.Security.Policy.Evidence Copy(this System.Security.Policy.Evidence evidence)
        {
            MemoryStream serializationStream = new MemoryStream();
            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(serializationStream, evidence);
            serializationStream.Position = 0L;
            return (System.Security.Policy.Evidence)formatter.Deserialize(serializationStream);
        }

        public static string Encrypt(string cleanString)
        {
            return BitConverter.ToString(Helper.EncryptToData(cleanString));
        }

        public static byte[] EncryptToData(string cleanString)
        {
            Byte[] clearBytes = new UnicodeEncoding().GetBytes(cleanString);
            return ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(clearBytes);
        }

        #endregion

        #region Generics

        public static IEnumerable<TSource> NullableUnion<TSource>(this IEnumerable<TSource> first, IEnumerable<TSource> second)
        {
            return first == null ? second : second == null ? first : first.Union(second);
        }

        public static T Max<T>(T o1, T o2)
            where T : IComparable<T>
        {
            return o1.CompareTo(o2) >= 0 ? o1 : o2;
        }

        public static T Min<T>(T o1, T o2)
            where T : IComparable<T>
        {
            return o1.CompareTo(o2) <= 0 ? o1 : o2;
        }

        public static Dictionary<TKey, TValue> SafeMergeWith<TKey, TValue>(this Dictionary<TKey, TValue> org, Dictionary<TKey, TValue> ext)
        {
            if (org == null || org.Count <= 0)
                return ext != null ? ext : new Dictionary<TKey, TValue>(0);
            else if (ext == null || ext.Count <= 0)
                return org != null ? org : new Dictionary<TKey, TValue>(0);
            else
            {
                Dictionary<TKey, TValue> result = new Dictionary<TKey, TValue>(org);
                foreach (KeyValuePair<TKey, TValue> item in ext)
                    result.Add(item.Key, item.Value);
                return result;
            }
        }

        public static Dictionary<TKey, TValue> SafeMergeWith<TSource, TKey, TValue>(this Dictionary<TKey, TValue> org, IEnumerable<TSource> ext, Func<TSource, TKey> extKeySelector, Func<TSource, TValue> extValueSelector)
        {
            if (org == null || org.Count <= 0)
                return ext != null ? ext.ToDictionary(extKeySelector, extValueSelector) : new Dictionary<TKey, TValue>(0);
            else if (ext == null || ext.Count() <= 0)
                return org != null ? org : new Dictionary<TKey, TValue>(0);
            else
            {
                Dictionary<TKey, TValue> result = new Dictionary<TKey, TValue>(org);
                foreach (TSource extItem in ext)
                    result.Add(extKeySelector(extItem), extValueSelector(extItem));
                return result;
            }
        }

        public static Dictionary<TKey, TValue> SafeMergeWith<TSource, TKey, TValue>(this IEnumerable<TSource> org, IEnumerable<TSource> ext, Func<TSource, TKey> extKeySelector, Func<TSource, TValue> extValueSelector)
        {
            if (org == null || org.Count() <= 0)
                return ext != null ? ext.ToDictionary(extKeySelector, extValueSelector) : new Dictionary<TKey, TValue>(0);
            else if (ext == null || ext.Count() <= 0)
                return org != null ? org.ToDictionary(extKeySelector, extValueSelector) : new Dictionary<TKey, TValue>(0);
            else
            {
                Dictionary<TKey, TValue> result = new Dictionary<TKey, TValue>(org.Count() + ext.Count());
                foreach (TSource extItem in org.Union(ext))
                    result.Add(extKeySelector(extItem), extValueSelector(extItem));
                return result;
            }
        }

        #endregion

        #region HEXA & BITS

        private readonly static char[] HEXA = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f' };
        private static int HEX2I(char c)
        {
            return c >= 'a' ? (c - 'a' + 10) : (c - '0');
        }

        public static char[] ReverseHex(char[] no)
        {
            char tmp;
            for (int i = 0, l = no.Length, n = l / 2; i < n; i++)
            {
                tmp = HEXA[0x0F - HEX2I(no[i])];
                no[i] = HEXA[0x0F - HEX2I(no[l - i - 1])];
                no[l - i - 1] = tmp;
            }

            return no;
        }

        /// <summary>
        /// First item has higher prioriy (if null will delegate to the second)
        /// </summary>
        public static short CopyBits(int? nullableBits, short nonnullBits)
        {
            if (nullableBits.HasValue)
                for (int i = 0; i < 16; i++)
                    switch ((nullableBits.Value >> (i * 2)) & 3)
                    {
                        case 2:
                            nonnullBits &= (short)~(1 << i);
                            break;
                        case 3:
                            nonnullBits |= (short)(1 << i);
                            break;
                    }
            return nonnullBits;
        }

        #endregion

        #region Shell Open

        public static void OpenWeb(string url, bool needHttp)
        {
            Helper.ShellOpen(!needHttp ? url : (@"http://" + url));
        }

        public static int OpenIE(string url)
        {
            return Helper.OpenIE(url, false, false);
        }

        public static int OpenIE(string url, bool needHttp, bool waitForExit)
        {
            Process proc = new Process();
            proc.EnableRaisingEvents = false;
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            proc.StartInfo.FileName = "iexplore";
            proc.StartInfo.Arguments = !needHttp ? url : (@"http://" + url);
            if (!waitForExit)
            {
                try { proc.Start(); } catch { }
                return 0;
            }
            else
            {
                try { proc.Start(); proc.WaitForExit(); } catch { }
                return proc.ExitCode;
            }
        }

        public static void ShellOpen(string path, bool waitForExit = false, string args = null, string verb = "open")
        {
            Process process = new Process();
            process.StartInfo.FileName = path;
            if (args != null)
                process.StartInfo.Arguments = args;
            process.StartInfo.Verb = verb;
            process.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
            if (!waitForExit)
            {
                try { process.Start(); } catch { }
            }
            else
            {
                try { process.Start(); process.WaitForExit(); } catch { }
            }
        }

        public static bool ShellOpen(string name, Stream stream)
        {
            if (stream == null)
                return false;
            var tempPath = Helper.GetTempFilePath(name);
            try
            {
                using (var file = System.IO.File.Create(tempPath))
                    StreamHelper.Streaming(stream, file);
            }
            catch
            {
                return false;
            }
            Helper.ShellOpen(tempPath);
            return true;
        }

        #endregion

        #region WIN32 Network

        public static string GetLocalIP()
        {
            try
            {
                var ips = System.Net.Dns.GetHostAddresses(System.Net.Dns.GetHostName());
                var localIP = ips == null || ips.Length <= 0 ? null : ips.FirstOrDefault(ip => !System.Net.IPAddress.IsLoopback(ip));
                if (localIP != null)
                    return localIP.ToString();
            }
            catch
            {
            }
            return "127.0.0.1";
        }

        public static string GetTempFilePath(string name)
        {
            int i = 1;
            string ext = null;
            var path = Path.Combine(Helper.TEMP_PATH, name);
            while (File.Exists(path))
            {
                try
                {
                    File.Delete(path);
                }
                catch
                {
                    path = Path.Combine(
                        Helper.TEMP_PATH,
                        String.Format(
                            "{0} ({1}){2}",
                            (ext != null ? name : Path.GetFileNameWithoutExtension(name)),
                            ++i,
                            ext ?? (ext = Path.GetExtension(name))));
                }
            }
            return path;
        }

        public static string ResolveComputer(string settingName, string defaultName)
        {
            if (String.IsNullOrWhiteSpace(settingName))
                return defaultName;
            settingName = settingName.Trim();
            return (settingName == "." || settingName.ToLowerInvariant() == "(local)") ? "localhost" : settingName;
        }

        #endregion

        #region Win32 Time

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        internal struct FILE_TIME
        {
            internal int ftTimeLow;
            internal int ftTimeHigh;
        }

        [System.Security.SuppressUnmanagedCodeSecurity, System.Runtime.InteropServices.DllImport("kernel32.dll")]
        internal static extern void GetSystemTimeAsFileTime(ref FILE_TIME lpSystemTimeAsFileTime);

        public static System.Timers.ElapsedEventArgs CreateTimerEventArgs()
        {
            FILE_TIME lpSystemTimeAsFileTime = new FILE_TIME();
            Helper.GetSystemTimeAsFileTime(ref lpSystemTimeAsFileTime);
            return (System.Timers.ElapsedEventArgs)Activator.CreateInstance(typeof(System.Timers.ElapsedEventArgs), DynamicDelegates.ObjectBindingFlags, null, new object[] { lpSystemTimeAsFileTime.ftTimeLow, lpSystemTimeAsFileTime.ftTimeHigh }, null);
        }

        #endregion

        #region Form Path

        private const string FormPathProtocol = @"app://";
        private static readonly int FormPathProtocolLength = FormPathProtocol.Length;

        public static string ToFormPath(this Tuple<Type, long?> formPath)
        {
            var moduleName = formPath.Item1.Assembly.ManifestModule.Name;
            var moduleParts = moduleName.Split(Helper.DOTS);
            var actualName = (moduleParts != null && moduleParts.Length == 4 && moduleParts[3].ToLowerInvariant() == "dll") ? moduleParts[2] : moduleName;
            return String.Format(
                @"{0}{1}/{2}{3}",
                Helper.FormPathProtocol,
                actualName,
                formPath.Item1.Name,
                !formPath.Item2.HasValue ? String.Empty : String.Format(@"/{0}", formPath.Item2));
        }

        public static Tuple<Type, long?> ParseFormPath(string pathString)
        {
            if (!pathString.StartsWith(Helper.FormPathProtocol))
                return null;
            var nextSlash = pathString.IndexOf('/', Helper.FormPathProtocolLength);
            if (nextSlash < 0 || nextSlash >= (pathString.Length - 1))
                return null;
            try
            {
                var assembly = Assembly.LoadFrom(pathString.Substring(Helper.FormPathProtocolLength, nextSlash - Helper.FormPathProtocolLength));
                var lastSlash = pathString.IndexOf('/', nextSlash + 1);
                if (lastSlash < 0)
                    lastSlash = pathString.Length;
                var formType = assembly.GetType(pathString.Substring(nextSlash + 1, lastSlash - nextSlash - 1));
                if (formType == null)
                    return null;
                long id;
                if ((lastSlash >= (pathString.Length - 1)) || !Int64.TryParse(pathString.Substring(lastSlash + 1), out id))
                    return new Tuple<Type, long?>(formType, null);
                return new Tuple<Type, long?>(formType, id);
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
            catch (System.IO.FileLoadException)
            {
                return null;
            }
            catch (System.BadImageFormatException)
            {
                return null;
            }
        }

        #endregion

        #region IO File Folder

        /// <summary>
        /// The build the map of FileName => FileInfo(fi.FullName, fi.Attributes, fi.CreationTimeUtc, fi.LastAccessTimeUtc, fi.LastWriteTimeUtc)
        /// </summary>
        /// <param name="searchPatterns">List of search parterns separated by '|', normal parterns are including, items with negative sign are excluding</param>
        public static Dictionary<string, Tuple<string, int, DateTime, DateTime, DateTime>> SearchFileInfos(string path, string searchPatterns)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            else
            {
                var patterns = searchPatterns.Split('|').Where(sp => !String.IsNullOrWhiteSpace(sp)).ToArray();
                if (patterns.Length > 0)
                {
                    var prefixLength = path.EndsWith(Path.DirectorySeparatorChar.ToString()) ? path.Length : (path.Length + 1);
                    return patterns
                        .Where(sp => sp[0] != '-')
                        .SelectMany(sp => new DirectoryInfo(path).GetFiles(sp, SearchOption.AllDirectories).Select(fi => fi.FullName))
                        .Except(patterns
                            .Where(sp => sp[0] == '-')
                            .SelectMany(sp => new DirectoryInfo(path).GetFiles(sp.Substring(1), SearchOption.AllDirectories).Select(fi => fi.FullName))
                        )
                        .Distinct()
                        .Select(f => new FileInfo(f))
                        .ToDictionary(fi => fi.FullName.Substring(prefixLength), fi => Helper.SearchFileInfoGet(fi));
                }
            }
            return new Dictionary<string, Tuple<string, int, DateTime, DateTime, DateTime>>();
        }

        public static bool SearchFileInfoDiffer(Dictionary<string, Tuple<string, int, DateTime, DateTime, DateTime>> files1, Dictionary<string, Tuple<string, int, DateTime, DateTime, DateTime>> files2)
        {
            return (files1.Count != files2.Count) ||
                    (
                        (files1.Count != 0) && (
                            files1.Keys.Except(files2.Keys).Any() ||
                            files1.Any(i => Helper.SearchFileInfoDiffer(i.Value, files2[i.Key]))
                        )
                    );
        }

        public static bool SearchFileInfoDiffer(Tuple<string, int, DateTime, DateTime, DateTime> file1, Tuple<string, int, DateTime, DateTime, DateTime> file2)
        {
            return file1.Item2 != file2.Item2 || file1.Item5 != file2.Item5;
        }

        public static Tuple<string, int, DateTime, DateTime, DateTime> SearchFileInfoGet(FileInfo fi)
        {
            return new Tuple<string, int, DateTime, DateTime, DateTime>(fi.FullName, (int)fi.Attributes, fi.CreationTimeUtc, fi.LastAccessTimeUtc, fi.LastWriteTimeUtc);
        }

        public static void SearchFileInfoSet(string fullpath, Tuple<string, int, DateTime, DateTime, DateTime> info)
        {
            System.IO.File.SetAttributes(fullpath, (FileAttributes)info.Item2);
            System.IO.File.SetCreationTimeUtc(fullpath, info.Item3);
            System.IO.File.SetLastAccessTimeUtc(fullpath, info.Item4);
            System.IO.File.SetLastWriteTimeUtc(fullpath, info.Item5);
        }

        public static string GetAppSystemName(string binPath)
        {
            var entryAsm = System.Reflection.Assembly.GetEntryAssembly();
            string moduleName;
            if (entryAsm != null)
                moduleName = entryAsm.ManifestModule.Name;
            else
            {
#if DEBUG
                var execs = (Helper.SearchFileName(binPath, "*.exe") ?? Helper.EMPTY_STRINGS).Where(p => !p.EndsWith(".vshost.exe")).ToArray();
#else
                var execs = Helper.SearchFileName(binPath, "*.exe");
#endif
                if (execs != null && execs.Length > 0)
                    moduleName = execs.First();
                else
                    moduleName = System.Reflection.Assembly.GetExecutingAssembly().ManifestModule.Name;
            }
            var _SYSTEM = System.IO.Path.GetFileNameWithoutExtension(moduleName);
#if DEBUG
            return _SYSTEM == "OneTest" ? "OneMES" : _SYSTEM;
#else
            return _SYSTEM;
#endif
        }

        public static string GetWebSystemName(string relPath)
        {
            var filenames = Helper.SearchFileName(relPath, "*.application");
            return filenames == null ? null : System.IO.Path.GetFileNameWithoutExtension(filenames.Single());
        }

        /// <summary>
        /// Help winform application to have App_Data (as WebForm)
        /// For example: Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\CallCenter.mdf;Integrated Security=True;User Instance=True
        /// </summary>
        public static void SetWinformDataDir()
        {
            string homeDir = AppDomain.CurrentDomain.BaseDirectory;
            if (homeDir.EndsWith(@"\bin\Debug\") || homeDir.EndsWith(@"\bin\Release\"))
                homeDir = System.IO.Directory.GetParent(homeDir).Parent.Parent.FullName;
            AppDomain.CurrentDomain.SetData("DataDirectory", System.IO.Path.Combine(homeDir, "App_Data"));
        }

        public static HashSet<string> SearchFileNames(string path, string pattern)
        {
            DirectoryInfo di = new DirectoryInfo(path);
            FileInfo[] files = di.GetFiles(pattern, SearchOption.TopDirectoryOnly);

            if (files == null || files.Length <= 0)
                return null;

            return new HashSet<string>(files.Select(o => Path.GetFileNameWithoutExtension(o.Name)));
        }

        public static HashSet<string> SearchDirNames(string path, string pattern)
        {
            DirectoryInfo di = new DirectoryInfo(path);
            var dirs = di.GetDirectories(pattern, SearchOption.TopDirectoryOnly);

            if (dirs == null || dirs.Length <= 0)
                return null;

            return new HashSet<string>(dirs.Select(o => Path.GetFileName(o.Name)));
        }

        public static string GetFullPath(string fromPath, string toPath)
        {
            return Path.GetFullPath(Path.IsPathRooted(toPath) ? toPath : Path.Combine(fromPath, toPath));
        }

        public static String[] SearchFileName(String path, String pattern)
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(path);
                if (!di.Exists)
                    return null;
                FileInfo[] files = di.GetFiles(pattern, SearchOption.TopDirectoryOnly);

                if (files == null || files.Length <= 0)
                    return null;

                return files.Select(o => o.Name).ToArray();
            }
            catch (IOException)
            {
                return null;
            }
        }

        public static IEnumerable<String> SearchFilePath(String path, String pattern)
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(path);
                FileInfo[] files = di.GetFiles(pattern, SearchOption.TopDirectoryOnly);

                if (files == null || files.Length <= 0)
                    return Helper.YieldEmpty<String>();

                return files.Select(o => o.FullName);
            }
            catch (IOException)
            {
                return Helper.YieldEmpty<String>();
            }
        }

        public static String[] SearchDirPathRecursive(String path, String pattern)
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(path);
                DirectoryInfo[] dirs = di.GetDirectories(pattern, SearchOption.AllDirectories);

                if (dirs == null || dirs.Length <= 0)
                    return null;

                return dirs.Select(o => o.FullName).ToArray();
            }
            catch (IOException)
            {
                return null;
            }
        }

        public static IEnumerable<string> SearchFilePathRecursive(String path, String pattern)
        {
            try
            {
                return new DirectoryInfo(path)
                    .GetFiles(pattern, SearchOption.AllDirectories)
                    .Select(o => o.FullName);
            }
            catch (IOException)
            {
                return Helper.YieldEmpty<String>();
            }
        }

        #endregion

        #region ShowError

        /// <summary>
        /// Nếu text NULL thì error.Message sẽ được hiển thị, còn không thì sẽ sử dụng text nếu text không Empty
        /// </summary>
        public static void ShowError(Exception error, string caption, System.Windows.Forms.IWin32Window owner, string text)
        {
            ShowError(error, caption, owner, text, (w, t, c, bt, ic, df) => System.Windows.Forms.MessageBox.Show(w, t, c, bt, ic, df));
        }

        public static void ShowError(Exception error, string caption, System.Windows.Forms.IWin32Window owner, string text, Func<System.Windows.Forms.IWin32Window, string, string, System.Windows.Forms.MessageBoxButtons, System.Windows.Forms.MessageBoxIcon, System.Windows.Forms.MessageBoxDefaultButton, System.Windows.Forms.DialogResult> showfunc)
        {
            var message = text == null ? error.FindInnerException().Message : text.Length == 0 ? null : text;
            var description = String.IsNullOrEmpty(message) ?
                "Có lỗi xảy ra hệ thống không thể tiếp tục, có muốn hiển thị thông tin chi tiết của thông báo lỗi?" :
                (message + "\n\nChương trình không thể tiếp tục, có muốn hiển thị thông tin chi tiết của thông báo lỗi?");
            if (System.Windows.Forms.DialogResult.Yes == showfunc(
                owner,
                description,
                caption,
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Error,
                System.Windows.Forms.MessageBoxDefaultButton.Button2))
            {
                if (System.Windows.Forms.DialogResult.Yes == showfunc(
                        owner,
                        String.Format("{1}{0}\n\nCó muốn chép thông tin chi tiết này vào Clipboard?", error, text == null || text.Length <= 0 ? String.Empty : (text + ". ")),
                        String.Format("{0}: {1}", caption, "Chi tiết lỗi hệ thống"),
                        System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button2))
                {
                    try
                    {
                        System.Windows.Forms.Clipboard.SetText(text == null || text.Length <= 0 ? Helper.ToString(error) : (text + ". " + Helper.ToString(error)));
                    }
                    catch (System.Runtime.InteropServices.ExternalException) { }
                }
            }
        }

        public static void ShowError(Exception error, string caption, System.Windows.Forms.IWin32Window owner)
        {
            Helper.ShowError(error, caption, owner, null);
        }

        public static void ShowError(Exception error, string caption, string text)
        {
            Helper.ShowError(error, caption, null, text);
        }

        public static void ShowError(Exception error, string caption)
        {
            Helper.ShowError(error, caption, null, null);
        }

        public static string ToString(Exception error)
        {
            return
                error is ReflectionTypeLoadException ? Helper.ToString((ReflectionTypeLoadException)error) : error.ToString();
        }

        public static string ToMessage(Exception error)
        {
            return error is ReflectionTypeLoadException ? Helper.ToMessage((ReflectionTypeLoadException)error) : error.Message;
        }
        private readonly static Assembly IgnoreAssembly = null;


        public static string ToMessage(ReflectionTypeLoadException error)
        {
            if (error.LoaderExceptions == null || error.LoaderExceptions.Length <= 0)
                return error.Message;

            HashSet<string> modules = new HashSet<string>();
            for (var i = 0; i < error.LoaderExceptions.Length; i++)
            {
                var module = error.LoaderExceptions[i] is FileNotFoundException && !String.IsNullOrEmpty(((FileNotFoundException)error.LoaderExceptions[i]).FileName) ?
                    ((FileNotFoundException)error.LoaderExceptions[i]).FileName : error.LoaderExceptions[i].Message;
                if (!modules.Contains(module))
                    modules.Add(module);
            }
            return modules.Count == 1 ? String.Format("Lỗi tải thư viện: {0}", modules.Single()) : modules.Select((m, i) => new { m = m, i = i }).Aggregate(new StringBuilder("Lỗi tải các thư viện:").AppendLine(), (sb, o) => sb.AppendFormat("  {0}.{1}", o.i, o.m).AppendLine(), sb => sb.ToString(Environment.NewLine.Length, sb.Length - Environment.NewLine.Length));
        }

        public static string ToString(ReflectionTypeLoadException error)
        {
            if (error.LoaderExceptions == null || error.LoaderExceptions.Length <= 0)
                return error.ToString();

            HashSet<string> messages = new HashSet<string>();
            HashSet<string> files = new HashSet<string>();
            var sb = new StringBuilder();
            for (var i = 0; i < error.LoaderExceptions.Length; i++)
            {
                if (error.LoaderExceptions[i] is FileNotFoundException &&
                    !String.IsNullOrEmpty(((FileNotFoundException)error.LoaderExceptions[i]).FileName) &&
                    !String.IsNullOrEmpty(((FileNotFoundException)error.LoaderExceptions[i]).FusionLog) &&
                    !((FileNotFoundException)error.LoaderExceptions[i]).FusionLog.StartsWith("WRN: Assembly binding logging is turned OFF."))
                {
                    if (!files.Contains(((FileNotFoundException)error.LoaderExceptions[i]).FileName) &&
                        files.Add(((FileNotFoundException)error.LoaderExceptions[i]).FileName))
                    {
                        sb.AppendFormat("  ERROR[{0}] ----------------------------------------", i + 1).AppendLine();
                        sb.Append("   Message: ").AppendLine(error.LoaderExceptions[i].Message);
                        sb.Append("   File Name: ").AppendLine(((FileNotFoundException)error.LoaderExceptions[i]).FileName);
                        sb.Append("   Fusion Log: ").AppendLine(((FileNotFoundException)error.LoaderExceptions[i]).FusionLog);
                    }
                }
                else
                {
                    if (!messages.Contains(error.LoaderExceptions[i].Message) &&
                        messages.Add(error.LoaderExceptions[i].Message))
                        sb.AppendFormat("  ERROR[{0}] = {1}", i + 1, error.LoaderExceptions[i].Message).AppendLine();
                }
            }
            if (error.Types != null && error.Types.Length > 0)
            {
                var errorTypes = error.Types.Where(t => t != null && t.Assembly != Helper.IgnoreAssembly).ToArray();
                if (errorTypes.Length > 0)
                    for (var i = 0; i < errorTypes.Length; i++)
                        sb.AppendFormat("  TYPE[{0}] = {1}", i + 1, errorTypes[i]).AppendLine();
            }
            sb.AppendLine("  EXCEPTION ----------------------------------------")
                .Append("   ").AppendLine(error.ToString());
            return sb.ToString(2, sb.Length - 2);
        }

        public static bool IsMatch(Exception ex, params Tuple<Type, MethodInfo, string, string>[] specialExceptions)
        {
            foreach (var se in specialExceptions)
                if ((se.Item1.IsAbstract ? se.Item1.IsAssignableFrom(ex.GetType()) : ex.GetType() == se.Item1) &&
                    (se.Item2 == null || Object.ReferenceEquals(ex.TargetSite, se.Item2)) &&
                    (se.Item3 == null || ex.Message == se.Item3) &&
                    (se.Item4 == null || ex.StackTrace == se.Item4))
                    return true;
            return false;
        }


        #endregion

        #region Others

        /// <summary>
        /// Xác định phần tử này có được load lên hay không theo cấu hình mã khách hàng (LicenseeCode = ONENET, YBNL, PYDX,..).
        /// Đây là tập các bộ luật cách nhau bởi dấu phẩy ',' và dấu loại trừ '-' trước tên KH để loại trừ KH này.
        /// Riêng dấu '*' thể hiện toàn bộ các khách hàng.
        /// Vd1: 'THPS,PYDX': dành riêng cho 2 KH
        /// Vd2: ''/null/'*': dành cho toàn bộ các khách hàng
        /// Vd3: '-THPS,-PYDX': dành cho toàn bộ các khách hàng, loại trừ 2 khách hàng
        /// </summary>
        public static bool ContainRules(string license, string rules)
        {
            if (rules.IndexOf(',') < 0)
            {
                switch (rules[0])
                {
                    case '-':
                        return license != rules.Substring(1);
                    case '*':
                        return true;
                    default:
                        return license == rules;
                }
            }
            else
            {
                var licenses = rules.Split(Helper.COMMAS);
                if (licenses.Length <= 0)
                    return true;

                var exclusions = licenses.Where(l => l[0] == '-').Select(l => l.Substring(1)).ToArray();
                if (exclusions.Length <= 0)
                    return licenses.Any(l => l.Length <= 0 || l == "*") ||
                        licenses.Contains(license);
                if (exclusions.Contains(license))
                    return false;

                var inclusions = licenses.Where(l => l[0] != '-').ToArray();
                if (inclusions.Length <= 0)
                    return true;
                return inclusions.Any(l => l.Length <= 0 || l == "*") ||
                    inclusions.Contains(license);
            }
        }

        public static bool ArraysEqual(Array a1, Array a2)
        {
            if (a1 == a2)
                return true;

            if (a1 == null || a2 == null)
                return false;

            if (a1.Length != a2.Length)
                return false;

            if (a1.Length <= 0)
                return true;

            IList list1 = a1, list2 = a2; //error CS0305: Using the generic type 'System.Collections.Generic.IList<T>' requires '1' type arguments
            for (int i = 0; i < a1.Length; i++)
            {
                if (!Object.Equals(list1[i], list2[i])) //error CS0021: Cannot apply indexing with [] to an expression of type 'IList'(x2)
                    return false;
            }
            return true;
        }

        public static System.Drawing.Imaging.ImageFormat GetImageFormat(string ext)
        {
            switch (ext)
            {
                case @".bmp":
                    return System.Drawing.Imaging.ImageFormat.Bmp;

                case @".gif":
                    return System.Drawing.Imaging.ImageFormat.Gif;

                case @".ico":
                    return System.Drawing.Imaging.ImageFormat.Icon;

                case @".jpg":
                case @".jpeg":
                    return System.Drawing.Imaging.ImageFormat.Jpeg;

                case @".png":
                    return System.Drawing.Imaging.ImageFormat.Png;

                case @".tif":
                case @".tiff":
                    return System.Drawing.Imaging.ImageFormat.Tiff;

                case @".wmf":
                    return System.Drawing.Imaging.ImageFormat.Wmf;

                default:
                    return null;
            }
        }

        #endregion
    }
}
