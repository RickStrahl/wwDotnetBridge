﻿using System;
using System.Runtime.InteropServices;
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


        public object GetValue()
        {
            var bridge = new wwDotNetBridge();
            return bridge.GetProperty(this,"Value");
        }

        /// <summary>
        /// Returns the name of the type in the Value structure
        /// </summary>
        /// <returns></returns>
        public string GetTypeName()
        {
            if (this.Value == null)
                return "null";

            return Value.GetType().FullName;
        }

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
        /// Sets an UInt64 value which is not supported
        /// in Visual FoxPro
        /// </summary>
        /// <param name="val"></param>
        public void SetUInt64(object val)
        {
            // unbox numeric value
            Value = Convert.ToUInt64(val);
        }

        /// <summary>
        /// Sets an UInt64 value which is not supported
        /// in Visual FoxPro
        /// </summary>
        /// <param name="val"></param>
        public void SetUInt32(object val)
        {
            // unbox numeric value
            Value = Convert.ToUInt32(val);
        }

        /// <summary>
        /// Sets a Single Value on the 
        /// </summary>
        /// <param name="val"></param>
        public void SetSingle(object val)
        {
            Value = Convert.ToSingle(val);
        }

        /// <summary>
        /// Set a float value
        /// </summary>
        /// <param name="val"></param>
        public void SetFloat(object val)
        {
            SetSingle(val);
        }

        /// <summary>
        /// Sets a character value from a string or integer
        /// </summary>
        /// <param name="val"></param>
        public void SetChar(object val)
        {
            Value = Convert.ToChar(val);
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
            Value = tval;
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
        /// <param name="enumValue">full type and value name. Example: System.Windows.Forms.MessageBoxOptions.OK</param>
        public void SetEnum(string enumValue)
        {
            if (string.IsNullOrEmpty(enumValue))
                throw new ArgumentNullException("Enum type is not set.");
            
            int at = enumValue.LastIndexOf('.');
            
            if (at == -1)
                throw new ArgumentException("Invalid format for enum value.");

            string type = enumValue.Substring(0, at);
            string property = enumValue.Substring(at+1);

            var bridge = new wwDotNetBridge();
            Value = bridge.GetStaticProperty(type,property);
        }

        /// <summary>
        /// Assigns an enum value that is based on a numeric (flag) value
        /// or a combination of flag values.
        /// </summary>
        /// <param name="enumType">Enum type name (System.Windows.Forms.MessageBoxOptions)</param>
        /// <param name="enumValue">numeric flag value to set enum to</param>
        public void SetEnumFlag(string enumType, int enumValue)
        {
            if (string.IsNullOrEmpty(enumType))
                throw new ArgumentNullException("Enum type is not set.");

            var bridge = new wwDotNetBridge();
            Type type = bridge.GetTypeFromName(enumType);

            Value = Enum.ToObject(type, enumValue);
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
        /// Method that sets the Value property by fixing up any
        /// values based on GetProperty() rules. This means if you pass 
        /// an ComArray the raw array will be unpacked and stored for example.
        /// </summary>
        public void SetValue(object value)
        {
            object obj = this;
            wwDotNetBridge bridge = new wwDotNetBridge();
            bridge.SetProperty(obj,"Value",value);
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

        public void SetValueFromEnum(string enumType, string enumName)
        {
            wwDotNetBridge bridge = new wwDotNetBridge();
            Value = bridge.GetStaticProperty(enumType, enumName);
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
        /// Invokes a static method with the passed parameters and sets the Value property
        /// from the result value.
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="method"></param>
        /// <param name="parms"></param>
        public void SetValueFromInvokeStaticMethod(string typeName, string method, ComArray parms)
        {
            wwDotNetBridge bridge = new wwDotNetBridge();
            var list = new List<object>();
            if (parms.Instance != null)
                foreach (object item in parms.Instance as IEnumerable)
                    list.Add(item);

            Value = bridge.InvokeStaticMethod_Internal(typeName, method, list.ToArray());
        }

        /// <summary>
        /// Invokes a method on the <see cref="System.Convert">System.Convert</see> static class 
        /// to perform conversions that are supported by that object
        /// </summary>
        /// <param name="method">The Convert method name to call as a string</param>
        /// <param name="value">The Value to convert</param>
        public void SetValueFromSystemConvert(string method, object value)
        {
            // we want the raw value
            Value = ReflectionUtils.CallStaticMethod("System.Convert", method,value);
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
            if (parms != null && parms.Instance != null)
                foreach(object item in parms.Instance as IEnumerable)                
                    list.Add(item);
                
              SetValueFromCreateInstance_Internal(typeName, list.ToArray());
        }

        /// <summary>
        /// Sets the Value property from a CreateInstance call. Useful for
        /// value types that can't be passed back to FoxPro.
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="parms"></param>
        protected void SetValueFromCreateInstance_Internal(string typeName, object[] parms)
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
            if (value == null || value is DBNull)
                Value = Guid.NewGuid();
            else if (value is ComGuid)
                Value = ((ComGuid) value).Guid;
            else if (value is String)
            {
                Value = new Guid(value as string);
            }
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

        /// <summary>
        /// Returns the string value of the embedded Value object
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            if (Value == null)
                return "null";

            return Value.ToString();
        }
    }
}
