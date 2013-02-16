using System;
using System.Runtime.InteropServices;
using Westwind.Utilities;
using System.Reflection;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Collections;

namespace Westwind.WebConnection
{
    /// <summary>
    /// Class that converts to various .NET types when passed
    /// a FoxPro value.
    /// 
    /// This object can then be used as an input to try and
    /// force parameters to a specific .NET type
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("Westwind.WebConnection.ComValue")]
    public class ComValue
    {
        /// <summary>
        /// Internally this value is set by the various
        /// SetXXX methods. It's of type objcect but 
        /// set to the appropriate .NET subtype.
        /// </summary>
        public object Value
        {
            get { return _Value; }
            set { _Value = value; }
        }
        private object _Value = null;


        /// <summary>
        /// Sets a Short value which is not supported
        /// in Visual FoxPro
        /// </summary>
        /// <param name="val"></param>
        public void SetInt16(object val)
        {
            // unbox numeric value
            Value = Convert.ToInt16(val);            
        }

        /// <summary>
        /// Sets an Int64 value which is not supported
        /// in Visual FoxPro
        /// </summary>
        /// <param name="val"></param>
        public void SetInt64(object val)
        {
            // unbox numeric value
            Value = Convert.ToInt64(val);
        }

        /// <summary>
        /// Sets a Decimal value. This can actually 
        /// be done in FoxPro with CAST(val as Currency)
        /// </summary>
        /// <param name="val"></param>
        public void SetDecimal(object val)
        {
            // unbox numeric value
            Value = Convert.ToDecimal(val);
        }

        /// <summary>
        /// Sets a Long (64 bit) integer value
        /// </summary>
        /// <param name="val"></param>
        public void SetLong(object val)
        {                     
            long tval = Convert.ToInt64(val);
            val = tval;
        }

        /// <summary>
        /// Returns a byte value which is similar
        /// to Int16.
        /// </summary>
        /// <param name="value"></param>
        public void SetByte(object value)
        {
            Value = Convert.ToByte(value);
        }


        /// <summary>
        /// Assigns an enum value to the value structure. This 
        /// allows to pass enum values to methods and constructors
        /// to ensure that method signatures match properly
        /// </summary>
        /// <param name="enumString">full type and value name. Example: System.Windows.Forms.MessageBoxOptions.OK</param>
        public void SetEnum(string enumString)
        {
            if (string.IsNullOrEmpty(enumString))
                throw new ArgumentNullException("Enum value not set.");
            
            int at = enumString.LastIndexOf('.');
            
            if (at == -1)
                throw new ArgumentException("Invalid format for enum value.");

            string type = enumString.Substring(0, at);
            string property = enumString.Substring(at+1);

            var bridge = new wwDotNetBridge();
            Value = bridge.GetStaticProperty(type,property);
        }



        /// <summary>
        /// Allows setting of DbNull from FoxPro since DbNull is an inaccessible
        /// value type for FoxPro.
        /// </summary>
        public void SetDbNull()
        {
            Value = DBNull.Value;
        }

        /// <summary>
        /// Sets the Value property from a property retrieved from .NET
        /// Useful to transfer value in .NET that are marshalled incorrectly
        /// in FoxPro such as Enum values (that are marshalled as numbers)
        /// </summary>
        /// <param name="objectRef">An object reference to the base object</param>
        /// <param name="property">Name of the property</param>
        public void SetValueFromProperty(object objectRef,string property)
        {
            wwDotNetBridge bridge = new wwDotNetBridge();
            Value = bridge.GetProperty(objectRef,property);
        }

        /// <summary>
        /// Sets the value property from a static property retrieved from .NET.
        /// Useful to transfer value in .NET that are marshalled incorrectly
        /// in FoxPro such as Enum values (that are marshalled as numbers)
        /// </summary>
        /// <param name="typeName">Full type name as a string - can also be an Enum type</param>
        /// <param name="property">The static property name</param>
        public void SetValueFromStaticProperty(string typeName,string property)
        {
            wwDotNetBridge bridge = new wwDotNetBridge();
            Value = bridge.GetStaticProperty(typeName,property);
        }

        /// <summary>
        /// Sets the Value property from a method call that passes it's positional arguments
        /// as an array. This version accepts a ComArray directly so it can be called
        /// directly from FoxPro with a ComArray instance
        /// </summary>
        /// <param name="objectRef">Object instance</param>
        /// <param name="method">Method to call</param>
        /// <param name="parms">An array of the parameters passed (use ComArray and InvokeMethod)</param>
        public void SetValueFromInvokeMethod(object objectRef, string method, ComArray parms)
        {
            var list = new List<object>();
            if (parms.Instance != null)
                foreach (object item in parms.Instance as IEnumerable)
                    list.Add(item);

            SetValueFromInvokeMethod(objectRef, method, list.ToArray());            
        }

        /// <summary>
        /// Sets the Value property from a method call that passes it's positional arguments
        /// as an array.
        /// </summary>
        /// <param name="objectRef">Object instance</param>
        /// <param name="method">Method to call</param>
        /// <param name="parms">An array of the parameters passed (use ComArray and InvokeMethod)</param>
        public void SetValueFromInvokeMethod(object objectRef, string method, object[] parms)
        {
            wwDotNetBridge bridge = new wwDotNetBridge();
            Value = bridge.InvokeMethodWithParameterArray(objectRef, method, parms);
        }

        /// <summary>
        /// Sets the Value property from a CreateInstance call. Useful for
        /// value types that can't be passed back to FoxPro.
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="parms"></param>
        public void SetValueFromCreateInstance(string typeName, ComArray parms)
        {
            var list = new List<object>();
            if (parms.Instance != null)
                foreach(object item in parms.Instance as IEnumerable)                
                    list.Add(item);
                
              SetValueFromCreateInstance(typeName, list.ToArray());
        }

        /// <summary>
        /// Sets the Value property from a CreateInstance call. Useful for
        /// value types that can't be passed back to FoxPro.
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="parms"></param>
        public void SetValueFromCreateInstance(string typeName, object[] parms)
        {            
            wwDotNetBridge bridge = new wwDotNetBridge();            
            Value = bridge.CreateInstance_Internal(typeName, parms);
        }

        /// <summary>
        /// Sets value to a .NET Guid. Creates a GUID from 
        /// either ComGuid instance
        /// a string, or if null creates a new GUID.                
        /// </summary>
        /// <param name="value"></param>
        public void SetGuid(object value)
        {
            if (value == null)
                Value = Guid.NewGuid();
            else if (value is ComGuid)
                Value = ((ComGuid) value).Guid;
            else if (value is String)
                Value = new Guid(value as string);
        }

        /// <summary>
        /// Create a new Guid on the Value structure
        /// </summary>
        public void NewGuid()
        {
            Value = Guid.NewGuid();
        }

        /// <summary>
        /// Retrieves a GUID value as a string
        /// from the Value structure
        /// </summary>
        /// <returns></returns>
        public string GetGuid()
        {
            if (Value == null)
                return null;

            Guid guid = (Guid) Value;
            return guid.ToString();
        }
    }
}
