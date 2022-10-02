#region License
/*
 **************************************************************
 *  Author: Rick Strahl 
 *          (c) West Wind Technologies, 2009-2018   
 *          http://www.west-wind.com/
 * 
 * Created: 4/10/2009
 *
 * Permission is hereby granted, free of charge, to any person
 * obtaining a copy of this software and associated documentation
 * files (the "Software"), to deal in the Software without
 * restriction, including without limitation the rights to use,
 * copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the
 * Software is furnished to do so, subject to the following
 * conditions:
 * 
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
 * OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
 * HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
 * WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
 * FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
 * OTHER DEALINGS IN THE SOFTWARE.
 **************************************************************  
*/
#endregion

// comment for OpenSource version
// #define WestwindProduct

using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Data;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Runtime.CompilerServices;


#if WestwindProduct
using Newtonsoft.Json;
#endif



namespace Westwind.WebConnection
{
    /// <summary>
    /// The wwDotNetBridge class provides a host of COM support functions for Visual FoxPro.
    /// It allows you to host the .NET runtime without relying on COM interop to load types,
    /// rather it acts as a proxy for instantiation and other tasks. 
    /// 
    /// This library can be used itself to load .NET and types, or you can use it as a helper
    /// with COM interop in which case you have to instantiate it as a COM object and call
    /// its methods directly rather than using the FoxPro helper class.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("Westwind.wwDotNetBridge")]
    public class wwDotNetBridge
    {

        private static bool _firstLoad = true;

        /// <summary>
        /// Returns error information if the call fails
        /// </summary>
        public string ErrorMessage { get; set; } = "";

        public bool Error { get; set; }

        public Exception LastException { get; set; }

        public bool IsThrowOnErrorEnabled { get; set; }

        public wwDotNetBridge()
        {
            if (Environment.Version.Major >= 4)
            {
                LoadAssembly("System.Core");

                if (_firstLoad)
                {
                    if (!ServicePointManager.SecurityProtocol.HasFlag(SecurityProtocolType.Tls12))
                    {
                        // support for all SSL/TLS protocols
                        ServicePointManager.SecurityProtocol =
                            SecurityProtocolType.Tls11 |
                            SecurityProtocolType.Tls12 |
                            SecurityProtocolType.Tls;
                    }

                    _firstLoad = false;
                }
            }
        }


        #region LoadAssembly Routines

        /// <summary>
        /// Loads an assembly into the AppDomain by its fully qualified assembly name
        /// </summary>
        /// <param name="AssemblyName"></param>
        /// <returns></returns>
        public bool LoadAssembly(string AssemblyName)
        {
            // *** use some predefined names to keep things simpler
            AssemblyName = FixupAssemblyName(AssemblyName);
            Assembly ass = null;
            try
            {
                ass = Assembly.Load(AssemblyName);
            }
            catch (Exception ex)
            {
                LastException = ex;
                ErrorMessage = ex.Message;
                return false;
            }

            return true;
        }

        /// <summary>
        /// Loads an assembly into the AppDomain by a fully qualified assembly path
        /// </summary>
        /// <param name="AssemblyFileName"></param>
        /// <returns></returns>
        public bool LoadAssemblyFrom(string AssemblyFileName)
        {
            try
            {
                Assembly.LoadFrom(AssemblyFileName);
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null)
                {
                    LastException = ex.InnerException;
                    SetError(ex.InnerException.Message);
                }
                else
                {
                    LastException = ex;
                    SetError(ex.Message);
                }

                return false;
            }

            return true;
        }

        #endregion

        #region CreateInstance by type name only

        /// <summary>
        /// Creates a type reference from a given type name if the
        /// assembly is already loaded.
        /// </summary>
        /// <param name="TypeName"></param>
        /// <returns></returns>
        public object CreateInstance(string TypeName)
        {
            object obj = CreateInstance_Internal(TypeName);
            return obj;
        }

        /// <summary>
        /// Creates a type reference from a given type name if the
        /// assembly is already loaded.
        /// </summary>        
        public object CreateInstance_OneParm(string TypeName, object Parm1)
        {
            Parm1 = FixupParameter(Parm1);

            return CreateInstance_Internal(TypeName, Parm1);
        }

        /// <summary>
        /// Creates a type reference from a given type name if the
        /// assembly is already loaded.
        /// </summary>
        public object CreateInstance_TwoParms(string TypeName, object Parm1, object Parm2)
        {
            Parm1 = FixupParameter(Parm1);
            Parm2 = FixupParameter(Parm2);

            return CreateInstance_Internal(TypeName, Parm1, Parm2);
        }

        /// <summary>
        /// Creates a type reference from a given type name if the
        /// assembly is already loaded.
        /// </summary>
        public object CreateInstance_ThreeParms(string TypeName, object Parm1, object Parm2, object Parm3)
        {
            Parm1 = FixupParameter(Parm1);
            Parm2 = FixupParameter(Parm2);
            Parm3 = FixupParameter(Parm3);

            return CreateInstance_Internal(TypeName, Parm1, Parm2, Parm3);
        }

        public object CreateInstance_FourParms(string TypeName, object Parm1, object Parm2, object Parm3, object Parm4)
        {
            Parm1 = FixupParameter(Parm1);
            Parm2 = FixupParameter(Parm2);
            Parm3 = FixupParameter(Parm3);
            Parm4 = FixupParameter(Parm4);

            return CreateInstance_Internal(TypeName, Parm1, Parm2, Parm3, Parm4);
        }
        public object CreateInstance_FiveParms(string TypeName, object Parm1, object Parm2, object Parm3, object Parm4, object Parm5)
        {
            Parm1 = FixupParameter(Parm1);
            Parm2 = FixupParameter(Parm2);
            Parm3 = FixupParameter(Parm3);
            Parm4 = FixupParameter(Parm4);
            Parm5 = FixupParameter(Parm5);

            return CreateInstance_Internal(TypeName, Parm1, Parm2, Parm3, Parm4, Parm5);
        }

        /// <summary>
        /// Creates an instance of a class  based on its type name. Assumes that the type's
        /// assembly is already loaded.
        /// 
        /// Note this will be a little slower than the versions that work with assembly
        /// name specified because this code has to search for the type first rather
        /// than directly activating it.
        /// </summary>
        /// <param name="TypeName"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        protected internal object CreateInstance_Internal(string TypeName, params object[] args)
        {
            SetError();

            if (args != null && args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    args[i] = FixupParameter(args[i]);
                }
            }

            object instance = null;
            Type type = null;
            try
            {
                type = GetTypeFromName(TypeName);

                if (type == null)
                {
                    SetError("Type not loaded. Please load call LoadAssembly first.");
                    return null;
                }

                instance = type.Assembly.CreateInstance(TypeName, false, BindingFlags.Default, null, args, null, null);
            }
            catch (TargetInvocationException ex)
            {
                SetError(ex, true);
                return null;
            }
            catch (Exception ex)
            {
                SetError(ex, true);
                return null;
            }

            return FixupReturnValue(instance);
        }


        /// <summary>
        /// Routine that loads an assembly by its 'application assembly name' - unsigned
        /// assemblies must be visible via the .NET path (current path or BIN dir) and
        /// GAC assemblies must be referenced by their full assembly name.
        /// </summary>
        /// <param name="AssemblyName"></param>
        /// <param name="TypeName"></param>
        /// <returns></returns>
        protected object CreateInstance_Internal(string AssemblyName, string TypeName, params object[] args)
        {
            SetError();

            AssemblyName = FixupAssemblyName(AssemblyName);

            object instance = null;
            try
            {
                if (args == null || args.Length == 0)
                    instance = AppDomain.CurrentDomain.CreateInstanceAndUnwrap(AssemblyName, TypeName);
                else
                {
                    for (int i = 0; i < args.Length; i++)
                    {
                        args[i] = FixupParameter(args[i]);
                    }
                    instance = AppDomain.CurrentDomain.CreateInstanceAndUnwrap(AssemblyName, TypeName, false,
                        BindingFlags.Default, null, args, null, null);
                }
                    
            }
            catch (TargetInvocationException ex)
            {
                SetError(ex, true);
                return null;
            }
            catch (Exception ex)
            {
                SetError(ex, true);
                return null;
            }

            return FixupReturnValue(instance);
        }



        /// <summary>
        /// Creates an instance of a type on an existing property of another type
        /// </summary>
        /// <param name="instance">Parent Instance that contains the property to set</param>
        /// <param name="property">The property on the parent instance to set</param>
        /// <param name="TypeName">Full name of the type to create</param>
        /// <returns></returns>
        public bool CreateInstanceOnType(object instance, string property, string TypeName)
        {
            return CreateInstanceOnType_Internal(instance, property, TypeName);
        }


        /// <summary>
        /// Creates an instance of a type on an existing property of another type
        /// </summary>
        /// <param name="instance">Parent Instance that contains the property to set</param>
        /// <param name="property">The property on the parent instance to set</param>
        /// <param name="TypeName">Full name of the type to create</param>
        /// <param name="parm1"></param>
        /// <returns></returns>
        public bool CreateInstanceOnType_OneParm(object instance, string property, string TypeName, object parm1)
        {
            return CreateInstanceOnType_Internal(instance, property, TypeName, parm1);
        }

        /// <summary>
        /// Creates an instance of a type on an existing property of another type
        /// </summary>
        /// <param name="instance">Parent Instance that contains the property to set</param>
        /// <param name="property">The property on the parent instance to set</param>
        /// <param name="TypeName">Full name of the type to create</param>
        /// <param name="parm1"></param>
        /// <param name="parm2"></param>
        /// <returns></returns>
        public bool CreateInstanceOnType_TwoParms(object instance, string property, string TypeName, object parm1,
            object parm2)
        {
            return CreateInstanceOnType_Internal(instance, property, TypeName, parm1, parm2);
        }

        /// <summary>
        /// Creates an instance of a type on an existing property of another type
        /// </summary>
        /// <param name="instance">Parent Instance that contains the property to set</param>
        /// <param name="property">The property on the parent instance to set</param>
        /// <param name="TypeName">Full name of the type to create</param>
        /// <param name="parm1"></param>
        /// <param name="parm2"></param>
        /// <param name="parm3"></param>
        /// <returns></returns>
        public bool CreateInstanceOnType_ThreeParms(object instance, string property, string TypeName, object parm1,
            object parm2, object parm3)
        {
            return CreateInstanceOnType_Internal(instance, property, TypeName, parm1, parm2, parm3);
        }

        /// <summary>
        /// Creates an instance of a type on an existing property of another type
        /// </summary>
        /// <param name="instance">Parent Instance that contains the property to set</param>
        /// <param name="property">The property on the parent instance to set</param>
        /// <param name="TypeName">Full name of the type to create</param>
        /// <param name="parm1"></param>
        /// <param name="parm2"></param>
        /// <param name="parm3"></param>
        /// <param name="parm4"></param>
        /// <returns></returns>
        public bool CreateInstanceOnType_FourParms(object instance, string property, string TypeName, object parm1,
            object parm2, object parm3, object parm4)
        {
            return CreateInstanceOnType_Internal(instance, property, TypeName, parm1, parm2, parm3, parm4);
        }

        /// <summary>
        /// Creates an instance of a type on an existing property of another type
        /// </summary>
        /// <param name="instance">Parent Instance that contains the property to set</param>
        /// <param name="property">The property on the parent instance to set</param>
        /// <param name="TypeName">Full name of the type to create</param>/// <param name="parm1"></param>
        /// <param name="parm2"></param>
        /// <param name="parm3"></param>
        /// <param name="parm4"></param>
        /// <param name="parm5"></param>
        /// <returns></returns>
        public bool CreateInstanceOnType_FiveParms(object instance, string property, string TypeName, object parm1,
            object parm2, object parm3, object parm4, object parm5)
        {
            return CreateInstanceOnType_Internal(instance, property, TypeName, parm1, parm2, parm3, parm4, parm5);
        }

        protected bool CreateInstanceOnType_Internal(object instance, string property, string TypeName,
            params object[] args)
        {
            SetError();

            object newInstance = null;
            Type type = null;
            try
            {
                type = GetTypeFromName(TypeName);

                if (type == null)
                {
                    SetError("Type not loaded. Please load call LoadAssembly first.");
                    return false;
                }

                newInstance =
                    type.Assembly.CreateInstance(TypeName, false, BindingFlags.Default, null, args, null, null);
            }
            catch (Exception ex)
            {
                LastException = ex;
                SetError(ex.Message);
                return false;
            }

            try
            {
                ReflectionUtils.SetPropertyExCom(instance, property, newInstance);
            }
            catch (Exception ex)
            {
                LastException = ex;
                ErrorMessage = ex.Message;
                return false;
            }

            return true;
        }

        #endregion

        #region Load instance by assembly name and type

        /// <summary>
        /// Creates an instance from a file reference with a parameterless constructor
        /// </summary>
        /// <param name="AssemblyFileName"></param>
        /// <param name="TypeName"></param>
        /// <returns></returns>
        public object CreateAssemblyInstanceFromFile(string AssemblyFileName, string TypeName)
        {
            return CreateInstanceFromFile_Internal(AssemblyFileName, TypeName, null);
        }

        /// <summary>
        /// Creates an instance from a file reference with a 1 parameter constructor
        /// </summary>
        /// <param name="AssemblyFileName"></param>
        /// <param name="TypeName"></param>
        /// <returns></returns>
        public object CreateAssemblyInstanceFromFile_OneParm(string AssemblyFileName, string TypeName, object Parm1)
        {
            return CreateInstanceFromFile_Internal(AssemblyFileName, TypeName, Parm1);
        }

        /// <summary>
        /// Creates an instance from a file reference with a two parameter constructor
        /// </summary>
        /// <param name="AssemblyFileName"></param>
        /// <param name="TypeName"></param>
        /// <returns></returns>
        public object CreateAssemblyInstanceFromFile_TwoParms(string AssemblyFileName, string TypeName, object Parm1,
            object Parm2)
        {
            return CreateInstanceFromFile_Internal(AssemblyFileName, TypeName, Parm1, Parm2);
        }

        /// <summary>
        /// Creates an instance from a file reference with a two parameter constructor
        /// </summary>
        /// <param name="assemblyFileName"></param>
        /// <param name="typeName"></param>
        /// <param name="parm1"></param>
        /// <returns></returns>
        public object CreateAssemblyInstanceFromFile_ThreeParms(string assemblyFileName, string typeName, object parm1, object parm2, object parm3)
        {
            return CreateInstanceFromFile_Internal(assemblyFileName, typeName, parm1, parm2, parm3);
        }


        /// <summary>
        /// Routine that loads a class from an assembly file name specified.
        /// </summary>
        /// <param name="AssemblyFileName"></param>
        /// <param name="TypeName"></param>
        /// <returns></returns>
        protected object CreateInstanceFromFile_Internal(string AssemblyFileName, string TypeName, params object[] args)
        {
            ErrorMessage = string.Empty;
            LastException = null;

            object instance = null;

            if (args != null && args.Length > 0)
            {
                for (var index = 0; index < args.Length; index++)
                    args[index] = FixupParameter(args[index]);
            }
            
            try
            {
                if (args == null)
                    instance = AppDomain.CurrentDomain.CreateInstanceFromAndUnwrap(AssemblyFileName, TypeName);
                else
                    instance = AppDomain.CurrentDomain.CreateInstanceFromAndUnwrap(AssemblyFileName, TypeName, false,
                        BindingFlags.Default, null, args, null, null);
            }
            catch (Exception ex)
            {
                LastException = ex;
                if (ex.InnerException != null)
                {
                    LastException = ex.InnerException;
                    SetError(ex.InnerException.Message);
                }
                else
                    SetError(ex.Message);

                return null;
            }

            return FixupReturnValue(instance);
        }


        /// <summary>
        /// Creates a new instance from a file file based assembly refence. Requires full
        /// filename including extension and path.
        /// </summary>
        /// <param name="AssemblyFileName"></param>
        /// <param name="TypeName"></param>
        /// <returns></returns>
        public object CreateAssemblyInstance(string AssemblyFileName, string TypeName)
        {
            return CreateInstance_Internal(AssemblyFileName, TypeName, null);
        }

        /// <summary>
        /// Creates a new instance from a file file based assembly refence. Requires full
        /// filename including extension and path.
        /// </summary>
        /// <param name="AssemblyFileName"></param>
        /// <param name="TypeName"></param>
        /// <returns></returns>
        public object CreateAssemblyInstance_OneParm(string AssemblyFileName, string TypeName, object Parm1)
        {
            return CreateInstance_Internal(AssemblyFileName, TypeName, Parm1);
        }

        /// <summary>
        /// Creates a new instance from a file file based assembly refence. Requires full
        /// filename including extension and path.
        /// </summary>
        /// <param name="AssemblyFileName"></param>
        /// <param name="TypeName"></param>
        /// <returns></returns>
        public object CreateAssemblyInstance_TwoParms(string AssemblyFileName, string TypeName, object Parm1,
            object Parm2)
        {
            return CreateInstance_Internal(AssemblyFileName, TypeName, Parm1, Parm2);
        }

        #endregion

        #region static member access

        public object InvokeStaticMethod(string TypeName, string Method)
        {
            return InvokeStaticMethod_Internal(TypeName, Method);
        }


        public object InvokeStaticMethod_OneParm(string TypeName, string Method, object Parm1)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1);
        }


        public object InvokeStaticMethod_TwoParms(string TypeName, string Method, object Parm1, object Parm2)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1, Parm2);
        }

        public object InvokeStaticMethod_ThreeParms(string TypeName, string Method, object Parm1, object Parm2,
            object Parm3)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1, Parm2, Parm3);
        }

        public object InvokeStaticMethod_FourParms(string TypeName, string Method, object Parm1, object Parm2,
            object Parm3, object Parm4)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1, Parm2, Parm3, Parm4);
        }

        public object InvokeStaticMethod_FiveParms(string TypeName, string Method, object Parm1, object Parm2,
            object Parm3, object Parm4, object Parm5)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1, Parm2, Parm3, Parm4, Parm5);
        }

        public object InvokeStaticMethod_SixParms(string TypeName, string Method, object Parm1, object Parm2,
            object Parm3, object Parm4, object Parm5, object Parm6)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6);
        }

        public object InvokeStaticMethod_SevenParms(string TypeName, string Method, object Parm1, object Parm2,
            object Parm3, object Parm4, object Parm5, object Parm6, object Parm7)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7);
        }

        public object InvokeStaticMethod_EightParms(string TypeName, string Method, object Parm1, object Parm2,
            object Parm3, object Parm4, object Parm5, object Parm6, object Parm7, object Parm8)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7,
                Parm8);
        }

        public object InvokeStaticMethod_NineParms(string TypeName, string Method, object Parm1, object Parm2,
            object Parm3, object Parm4, object Parm5, object Parm6, object Parm7, object Parm8, object Parm9)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7, Parm8,
                Parm9);
        }

        public object InvokeStaticMethod_TenParms(string TypeName, string Method, object Parm1, object Parm2,
            object Parm3, object Parm4, object Parm5, object Parm6, object Parm7, object Parm8, object Parm9,
            object Parm10)
        {
            return InvokeStaticMethod_Internal(TypeName, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7, Parm8,
                Parm9, Parm10);
        }

        /// <summary>
        /// Invokes a static method
        /// </summary>
        /// <param name="TypeName"></param>
        /// <param name="Method"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        internal object InvokeStaticMethod_Internal(string TypeName, string Method, params object[] args)
        {
            SetError();

            Type type = GetTypeFromName(TypeName);
            if (type == null)
            {
                SetError("Type is not loaded. Please make sure you call LoadAssembly first.");
                return null;
            }

            // fix up all parameters
            object[] ar;
            if (args == null || args.Length == 0)
                ar = new object[0];
            else
            {
                ar = new object[args.Length];
                for (int i = 0; i < args.Length; i++)
                {
                    ar[i] = FixupParameter(args[i]);
                }
            }

            ErrorMessage = "";
            object Result = null;
            try
            {
                Result = type.InvokeMember(Method,
                    BindingFlags.Static | BindingFlags.Public | BindingFlags.InvokeMethod, null, type, ar);
            }
            catch (Exception ex)
            {
                var innerEx = ex.GetBaseException();
                SetError(innerEx, true);
                throw innerEx;
            }

            // Update ComValue parameters to support ByRef Parameters
            for (int i = 0; i < ar.Length; i++)
            {
                if (args[i] is ComValue)
                {
                    if(ar[i] is ComValue)
                        ((ComValue) args[i]).Value = ((ComValue) ar[i]).Value;
                    else
                        ((ComValue) args[i]).Value = FixupReturnValue(ar[i]);
                }
            }

            return FixupReturnValue(Result);
        }

        /// <summary>
        /// Retrieves a value from  a static property by specifying a type full name and property
        /// </summary>
        /// <param name="TypeName">Full type name (namespace.class)</param>
        /// <param name="Property">Property to get value from</param>
        /// <returns></returns>
        public object GetStaticProperty(string TypeName, string Property)
        {
            SetError();

            Type type = GetTypeFromName(TypeName);
            if (type == null)
            {
                SetError("Type is not loaded. Please make sure you call LoadAssembly first.");
                return null;
            }

            object val = null;
            try
            {
                val = type.InvokeMember(Property,
                    BindingFlags.Static | BindingFlags.Public | BindingFlags.GetField | BindingFlags.GetProperty, null,
                    type, null);
                val = FixupReturnValue(val);
            }
            catch (Exception ex)
            {
                SetError(ex.GetBaseException(), true);
                throw ex.GetBaseException();
            }

            return val;
        }

        public bool SetStaticProperty(string typeName, string property, object value)
        {
            SetError();


            Type type = GetTypeFromName(typeName);
            if (type == null)
            {
                SetError("Unable to find type signature. Please make sure you call LoadAssembly first.");
                return false;
            }

            value = FixupParameter(value);

            ErrorMessage = "";
            try
            {
                type.InvokeMember(property,
                    BindingFlags.Static | BindingFlags.Public | BindingFlags.SetField | BindingFlags.SetProperty, null,
                    type, new object[1] {value});
            }
            catch (Exception ex)
            {
                SetError(ex.GetBaseException(), true);
                throw ex.GetBaseException();
            }

            return true;
        }

        #endregion

        #region Instance Member access

        /// <summary>
        /// Invokes a method with no parameters
        /// </summary>
        /// <param name="Instance"></param>
        /// <param name="Method"></param>
        /// <returns></returns>
        public object InvokeMethod(object Instance, string Method)
        {
            return InvokeMethod_Internal(Instance, Method);
        }


        /// <summary>
        /// Invokes a method with an explicit array of parameters
        /// Allows for any number of parameters to be passed.
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="method"></param>
        /// <param name="parms"></param>
        /// <returns></returns>
        public object InvokeMethodWithParameterArray(object instance, string method, object[] parms)
        {
            return InvokeMethod_InternalWithObjectArray(instance, method, parms);
        }


        /// <summary>
        /// Invokes a method with one parameter
        /// </summary> 
        /// <param name="Instance"></param>
        /// <param name="Method"></param>
        /// <param name="Parm1"></param>
        /// <returns></returns>
        public object InvokeMethod_OneParm(object Instance, string Method, object Parm1)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1);
        }

        /// <summary>
        /// Invokes a method with two parameters
        /// </summary>
        /// <param name="Instance"></param>
        /// <param name="Method"></param>
        /// <param name="Parm1"></param>
        /// <param name="Parm2"></param>
        /// <returns></returns>
        public object InvokeMethod_TwoParms(object Instance, string Method, object Parm1, object Parm2)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2);
        }

        public object InvokeMethod_ThreeParms(object Instance, string Method, object Parm1, object Parm2, object Parm3)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3);
        }

        public object InvokeMethod_FourParms(object Instance, string Method, object Parm1, object Parm2, object Parm3,
            object Parm4)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3, Parm4);
        }

        public object InvokeMethod_FiveParms(object Instance, string Method, object Parm1, object Parm2, object Parm3,
            object Parm4, object Parm5)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3, Parm4, Parm5);
        }

        public object InvokeMethod_SixParms(object Instance, string Method, object Parm1, object Parm2, object Parm3,
            object Parm4, object Parm5, object Parm6)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6);
        }

        public object InvokeMethod_SevenParms(object Instance, string Method, object Parm1, object Parm2, object Parm3,
            object Parm4, object Parm5, object Parm6, object Parm7)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7);
        }

        public object InvokeMethod_EightParms(object Instance, string Method, object Parm1, object Parm2, object Parm3,
            object Parm4, object Parm5, object Parm6, object Parm7, object Parm8)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7, Parm8);
        }

        public object InvokeMethod_NineParms(object Instance, string Method, object Parm1, object Parm2, object Parm3,
            object Parm4, object Parm5, object Parm6, object Parm7, object Parm8, object Parm9)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7, Parm8,
                Parm9);
        }

        public object InvokeMethod_TenParms(object Instance, string Method, object Parm1, object Parm2, object Parm3,
            object Parm4, object Parm5, object Parm6, object Parm7, object Parm8, object Parm9, object Parm10)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7, Parm8,
                Parm9, Parm10);
        }

        public object InvokeMethod_ElevenParms(object Instance, string Method, object Parm1, object Parm2, object Parm3,
            object Parm4, object Parm5, object Parm6, object Parm7, object Parm8, object Parm9, object Parm10,
            object Parm11)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7, Parm8,
                Parm9, Parm10, Parm11);
        }

        public object InvokeMethod_TwelveParms(object Instance, string Method, object Parm1, object Parm2, object Parm3,
            object Parm4, object Parm5, object Parm6, object Parm7, object Parm8, object Parm9, object Parm10,
            object Parm11, object Parm12)
        {
            return InvokeMethod_Internal(Instance, Method, Parm1, Parm2, Parm3, Parm4, Parm5, Parm6, Parm7, Parm8,
                Parm9, Parm10, Parm11, Parm12);
        }

        public object InvokeMethod_ParameterArray(object Instance, string Method, object ParameterArray)
        {
            ComArray ParmArray = ParameterArray as ComArray;
            for (int i = 0; i < ParmArray.Count; i++)
            {
                Array ar = ParmArray.Instance as Array;
                ar.SetValue(FixupParameter(ar.GetValue(i)), i);
            }

            return InvokeMethod_InternalWithObjectArray(Instance, Method, ParmArray.Instance as object[]);
        }

        internal object InvokeMethod_Internal(object instance, string method, params object[] args)
        {
            var fixedInstance = instance;
            if (instance is ComValue)
                fixedInstance = ((ComValue) instance).Value;

            SetError();

            object[] ar;
            if (args == null || args.Length == 0)
                ar = new object[0];
            else
            {
                ar = new object[args.Length];
                for (int i = 0; i < args.Length; i++)
                {
                    ar[i] = FixupParameter(args[i]);
                }
            }

            object result = null;
            try
            {
                if (method.Contains(".") || method.Contains("["))
                    result = ReflectionUtils.CallMethodEx(fixedInstance, method, ar);
                else
                    result = ReflectionUtils.CallMethodCom(fixedInstance, method, ar);
            }
            catch (Exception ex)
            {
                SetError(ex.GetBaseException(), true);
                throw ex.GetBaseException();
            }

            // Update ComValue parameters to support ByRef Parameters
            for (int i = 0; i < ar.Length; i++)
            {
                if (args[i] is ComValue)
                {
                    if(ar[i] is ComValue)
                        ((ComValue) args[i]).Value = ((ComValue) ar[i]).Value;
                    else
                        ((ComValue) args[i]).Value = FixupReturnValue(ar[i]);
                }
            }

            return FixupReturnValue(result);
        }



        protected object InvokeMethod_InternalWithObjectArray(object instance, string method, object[] args)
        {

            var fixedInstance = instance;
            if (fixedInstance is ComValue)
                fixedInstance = ((ComValue) instance).Value;

            SetError();
            object result;

            try
            {
                if (method.Contains(".") || method.Contains("["))
                    result = ReflectionUtils.CallMethodEx(fixedInstance, method, args);
                else
                    result = ReflectionUtils.CallMethodCom(fixedInstance, method, args);
            }
            catch (Exception ex)
            {
                SetError(ex.GetBaseException(), true);
                throw ex.GetBaseException();
            }

            // Update ComValue parameters to support ByRef Parameters
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] is ComValue)
                {
                    if(args[i] is ComValue)
                        ((ComValue) args[i]).Value = ((ComValue) args[i]).Value;
                    else
                        ((ComValue) args[i]).Value = FixupReturnValue(args[i]);
                }
            }

            return FixupReturnValue(result);
        }


        public object GetProperty(object instance, string property)
        {
            LastException = null;
            try
            {
                object fixedInstance = null;
                if (instance is ComValue)
                    fixedInstance = ((ComValue) instance).Value;
                else 
                    fixedInstance = instance;

                object val;

                if (property.Contains(".") || property.Contains("["))
                    val = ReflectionUtils.GetPropertyEx(fixedInstance, property);
                else
                    val = ReflectionUtils.GetPropertyCom(fixedInstance, property);  // this is more reliable in that it handles value conversions

                val = FixupReturnValue(val);
                return val;
            }
            catch (Exception ex)
            {
                SetError(ex.GetBaseException(), true);
                throw ex.GetBaseException();
            }
        }


        /// <summary>
        /// Returns a property value by allowing . syntax to drill
        /// into nested objects. Use this method to step over objects 
        /// that FoxPro can't directly access (like structs, generics etc.)
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="property"></param>
        /// <returns></returns>
        public object GetPropertyEx(object instance, string property)
        {
            var fixedInstance = instance;
            if (fixedInstance is ComValue)
                fixedInstance = ((ComValue) fixedInstance).Value;

            LastException = null;
            try
            {
                object val = ReflectionUtils.GetPropertyEx(fixedInstance, property);
                val = FixupReturnValue(val);
                return val;
            }
            catch (Exception ex)
            {
                SetError(ex.GetBaseException(), true);
                throw ex.GetBaseException();
            }
        }

        /// <summary>
        /// Sets a property of a .NET object with a value
        /// </summary>
        /// <param name="Instance"></param>
        /// <param name="property"></param>
        /// <param name="value"></param>
        public void SetProperty(object instance, string property, object value)
        {

            var fixedInstance = instance;
            if (fixedInstance is ComValue)
                fixedInstance = ((ComValue) fixedInstance).Value;

            LastException = null;

            if (value is DBNull)
                value = null;

            value = FixupParameter(value);
            try
            {
                if (property.Contains(".") || property.Contains("["))
                    ReflectionUtils.SetPropertyEx(fixedInstance, property, value);
                else
                    ReflectionUtils.SetPropertyCom(fixedInstance, property, value);
            }
            catch (Exception ex)
            {
                SetError(ex.GetBaseException(), true);
                throw ex.GetBaseException();
            }
        }

        /// <summary>
        /// Sets a property of a .NET object with a value using extended syntax.
        /// 
        /// This method supports '.' syntax so you can use "Property.ChildProperty"
        /// to walk the object hierarchy in the string property parameter. 
        /// 
        /// This method also supports accessing of Array/Collection indexers (Item[1])
        /// </summary>
        /// <param name="Instance"></param>
        /// <param name="Property"></param>
        /// <param name="Value"></param>
        public void SetPropertyEx(object instance, string Property, object Value)
        {

            var fixedInstance = instance;
            if (fixedInstance is ComValue)
                fixedInstance = ((ComValue) fixedInstance).Value;

            LastException = null;
            try
            {
                if (Value is DBNull)
                    Value = null;

                Value = FixupParameter(Value);
                ReflectionUtils.SetPropertyEx(fixedInstance, Property, Value);
            }
            catch (Exception ex)
            {
                LastException = ex;
                SetError(ex.GetBaseException(), true);
                throw ex.GetBaseException();
            }
        }

        /// <summary>
        /// Returns the name of an enum field given an enum value
        /// passed. Pass in the name of the enum type
        /// </summary>
        /// <param name="EnumTypeName"></param>
        /// <param name="Value"></param>
        /// <returns></returns>
        public string GetEnumString(string EnumTypeName, object Value)
        {
            SetError();
            try
            {
                Type type = GetTypeFromName(EnumTypeName);
                return Enum.GetName(type, Value);
            }
            catch (Exception ex)
            {
                SetError(ex.GetBaseException());
                throw ex.GetBaseException();
            }
        }

        #endregion

        #region Async Methods

        /// <summary>
        /// Invokes a method on asynchronously and fires OnCompleted and OnError
        /// events on a passed in callback object.
        /// </summary>
        /// <param name="callBack">
        /// A callback object that has to have two methods:
        /// OnCompleted(lvResult, lcMethod)
        /// OnError(lcErrorMsg,loException, lcMethod)        
        /// </param>
        /// <param name="instance"></param>
        /// <param name="method"></param>
        /// <param name="parameters"></param>
        public void InvokeStaticMethodAsync(object callBack, string typeName, string method, params object[] parameters)
        {
            if (callBack == null || string.IsNullOrEmpty(method))
                throw new ApplicationException("You have to pass a callback object and method name.");

            var parms = new object[5];
            parms[0] = callBack;
            parms[1] = typeName;
            parms[2] = method;
            parms[3] = parameters;
            parms[4] = true; // isStatic

            //var t = new Thread(_InvokeMethodAsync);
            //t.Start(parms);

            Task.Run(() => _InvokeMethodAsync(parms));
        }

        /// <summary>
        /// Invokes a method on a new thread and fires OnCompleted and OnError
        /// events on a passed in callback object.
        /// </summary>
        /// <param name="callBack">
        /// A callback object that has to have two methods:
        /// OnCompleted(lvResult, lcMethod)
        /// OnError(lcErrorMsg,loException, lcMethod)        
        /// </param>
        /// <param name="instance"></param>
        /// <param name="method"></param>
        /// <param name="parameters"></param>
        public void InvokeMethodAsync(object callBack, object instance, string method, params object[] parameters)
        {
            if (callBack == null || string.IsNullOrEmpty(method))
                throw new ApplicationException("You have to pass a callback object and method name.");

            var parms = new object[5];
            parms[0] = callBack;
            parms[1] = instance;
            parms[2] = method;
            parms[3] = parameters;
            parms[4] = false; // isStatic

            //var t = new Thread(_InvokeMethodAsync);
            //t.Start(parms);

            Task.Run(() => _InvokeMethodAsync(parms));
        }


        public void InvokeTaskMethodAsync(
            object callBack, 
            object instance, 
            string method,
            params object[] parameters)
        {
            bool isStatic = instance is string;
            if (callBack is DBNull)
                callBack = null;
            
            Task result = null;
            try
            {
                if (!isStatic)
                {
                    if (parameters == null || parameters.Length < 1)
                        result = InvokeMethod_Internal(instance, method) as Task;
                    else
                        result = InvokeMethod_Internal(instance, method, parameters) as Task;
                }
                else
                {
                    if (parameters == null || parameters.Length < 1)
                        result = InvokeStaticMethod_Internal(instance as string, method) as Task;
                    else
                        result = InvokeStaticMethod_Internal(instance as string, method, parameters) as Task;
                }
            }
            catch (Exception ex)
            {
                if (callBack != null)
                {
                    try
                    {
                        InvokeMethod_Internal(callBack, "onError", ex.Message, ex.GetBaseException(), method);
                    }
                    catch
                    {
                        // no error method - just eat it
                        LastException = ex;
                    }
                }

                return;
            }

            if (callBack != null)
            {
                try
                {
                    result.ContinueWith((r) =>
                    {
                        object res = null;
                        var t = result.GetType();

                        if (t.IsGenericType)
                        {
                            res = GetProperty(result, "Result");
                        }
                        InvokeMethod_Internal(callBack, "onCompleted", res, method);
                    });
                }
                catch (Exception ex)
                {
                    // no callback method - just eat it
                    LastException = ex;
                }
            }
        }


        /// <summary>
        /// Internal handler method that actually makes the async call on a thread
        /// </summary>
        /// <param name="parmList"></param>
        private void _InvokeMethodAsync(object parmList)
        {
            object[] parms = parmList as object[];

            object callBack = parms[0];
            if (callBack is DBNull)
                callBack = null;

            object instance = parms[1];
            string method = parms[2] as string;
            object[] parameters = parms[3] as object[];

            bool isStatic = false;
            if (parms.Length > 4)
                isStatic = (bool) parms[4];

            object result = null;
            try
            {
                if (!isStatic)
                {
                    if (parameters == null || parameters.Length < 1)
                        result = InvokeMethod_Internal(instance, method);
                    else
                        result = InvokeMethod_Internal(instance, method, parameters);
                }
                else
                {
                    if (parameters == null || parameters.Length < 1)
                        result = InvokeStaticMethod_Internal(instance as string, method);
                    else
                        result = InvokeStaticMethod_Internal(instance as string, method, parameters);
                }

            }
            catch (Exception ex)
            {
                if (callBack != null)
                {
                    try
                    {
                        InvokeMethod_Internal(callBack, "onError", ex.Message, ex.GetBaseException(), method);
                    }
                    catch
                    {
                        // no error method - just eat it
                        LastException = ex;
                    }

                }

                return;
            }

            if (callBack != null)
            {
                try
                {
                    InvokeMethod_Internal(callBack, "onCompleted", result, method);
                }
                catch (Exception ex)
                {
                    // no callback method - just eat it
                    LastException = ex;
                }
            }
        }

        #endregion

        #region Array Functions

        /// <summary>
        /// Creates an instance of an array on a given base object instance by name.
        /// Array is created with 'empty' elements - ie. objects are null and value
        /// types are set to their default() values.
        /// </summary>
        /// <param name="baseType"></param>
        /// <param name="arrayProperty"></param>
        /// <returns></returns>
        private Array CreateArrayInstanceInternal(object baseType, string arrayProperty, int size)
        {
            MemberInfo[] miArray = baseType.GetType().GetMember(arrayProperty);
            if (miArray == null || miArray.Length == 0)
            {
                SetError("Array member doesn't exist on base type");
                return null;
            }

            Type type = null;
            MemberInfo mi = miArray[0];
            if (mi.MemberType == MemberTypes.Field)
                type = baseType.GetType().GetField(arrayProperty).FieldType.GetElementType();
            else
                type = baseType.GetType().GetProperty(arrayProperty).PropertyType.GetElementType();

            if (type == null)
            {
                SetError("Invalid type for array: " + type.Name);
                return null;
            }

            // *** Create instance and assign size
            Array ar = Array.CreateInstance(type, size);
            return ar;
        }

        /// <summary>
        /// Creates an array instance of a given type and size. Note the
        /// elements of this array are null/default and need to be set explicitly
        /// </summary>
        /// <param name="baseType">Object instance on which to create the array</param>
        /// <param name="arrayProperty">String property/field name of the array to create</param>
        /// <param name="size">Size of the array to createArray</param>
        /// <returns></returns>
        public bool CreateArrayOnInstance(object baseType, string arrayProperty, int size)
        {
            SetError();

            Array ar = CreateArrayInstanceInternal(baseType, arrayProperty, size);
            if (ar == null)
                return false;

            ReflectionUtils.SetPropertyExCom(baseType, arrayProperty, ar);

            return true;
        }

        /// <summary>
        /// Creates a new array instance on a type of exactly 1 array item which is
        /// assigned the item parameter passed in. 
        /// </summary>
        /// <param name="baseType"></param>
        /// <param name="arrayProperty"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool CreateArrayOnInstanceWithObject(object baseType, string arrayProperty, object item)
        {
            SetError();

            // *** Create instance and assign
            Array ar = CreateArrayInstanceInternal(baseType, arrayProperty, 1);
            if (ar == null)
                return false;

            // *** assign the item passed in
            ar.SetValue(item, 0);

            ReflectionUtils.SetPropertyExCom(baseType, arrayProperty, ar);

            return true;
        }

        /// <summary>
        /// Creates an instance of an array
        /// </summary>
        /// <param name="arrayTypeString"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        public ComArray CreateArray(string arrayTypeString)
        {
            SetError();

            Type type = null;
            type = ReflectionUtils.GetTypeFromName(arrayTypeString);

            if (type == null)
            {
                SetError("Invalid type for array: " + arrayTypeString);
                return null;
            }

            ComArray comArray = new ComArray();

            // *** Create instance and assign
            Array ar = Array.CreateInstance(type, 0);

            // *** assign the item passed in
            //ar.SetValue(item, 0);
            comArray.Instance = ar;

            return comArray;
        }

        /// <summary>
        /// Creates an array from a specific instance of a COM object
        /// </summary>
        /// <param name="instance"></param>
        /// <returns></returns>
        public ComArray CreateArrayFromInstance(object instance)
        {
            return new ComArray(instance);
        }


        /// <summary>
        /// Used to add an item to an array by indirection to work around VFP's
        /// inability to easily add array elements.
        /// </summary>
        /// <param name="baseObject">The object that has the Array property</param>
        /// <param name="arrayObject">The array property name as a string</param>
        /// <param name="item">The item to set it to. Should not be null.</param>
        /// <returns></returns>
        public bool AddArrayItem(object baseObject, string arrayObject, object item)
        {
            SetError();

            Array ar = baseObject.GetType().InvokeMember(arrayObject,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.GetField,
                null, baseObject, null) as Array;

            // *** Null assignments are not allowed because we may have to create the array
            // *** and a type is required for that. If necessary create an empty instance
            if (item == null && ar == null)
                return false;


            // *** This may be ambigous - could mean no property or array exists and is null
            if (ar == null)
            {
                ar = Array.CreateInstance(item.GetType(), 1);
                if (ar == null)
                    return false;

                ar.SetValue(item, 0);
            }
            else
            {
                Type itemType = ar.GetType().GetElementType();

                int Size = ar.GetLength(0); // one dimensional

                // resize the array by creating a new one
                Array TempArray = Array.CreateInstance(itemType, Size + 1);
                ar.CopyTo(TempArray, 0);

                // reassign to original instance var
                ar = TempArray;

                // and assign the new added item
                ar.SetValue(item, Size);
            }

            // make sure the original array instance reference gets updated
            // both for new array and updated array!
            baseObject.GetType().InvokeMember(arrayObject,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty | BindingFlags.SetField,
                null, baseObject, new object[1] {ar});

            return true;
        }


        private object _tvalue;

        /// <summary>
        /// Returns an indexed property Value
        /// </summary>
        /// <param name="baseList">List object</param>
        /// <param name="index">Index into the list</param>
        /// <returns></returns>
        public object GetIndexedProperty(object baseList, int index)
        {
            try
            {
                _tvalue = baseList;
                return GetPropertyEx(this, "_tvalue[" + index + "]");
            }
            finally
            {
                _tvalue = null;
            }
        }


        /// <summary>
        /// Returns an individual Array Item by its index
        /// </summary>
        /// <param name="baseObject"></param>
        /// <param name="arrayName"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public object GetArrayItem(object baseObject, string arrayName, int index)
        {
            SetError();

            Array ar = baseObject.GetType().InvokeMember(arrayName,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.GetField,
                null, baseObject, null) as Array;

            // *** Null assignments are not allowed because we may have to create the array
            // *** and a type is required for that. If necessary create an empty instance
            if (ar == null)
                return false;

            return ar.GetValue(index);
        }
        
        
        /// <summary>
        /// Returns a dictionary item from a dictionary object by passing in a key.
        /// </summary>
        /// <param name="dictionary"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public object GetDictionaryItem(IDictionary dictionary, object key)
        {
            object result = dictionary[key];
            return result;
        }


        /// <summary>
        /// Sets an array element to a given value. Assumes the array is big
        /// enough and the array item exists.
        /// </summary>
        /// <param name="baseObject">base object reference</param>
        /// <param name="arrayName">Name of the array as a string</param>
        /// <param name="index">The index of the item to set</param>
        /// <param name="value">The value to set the array item to</param>
        /// <returns></returns>
        public bool SetArrayItem(object baseObject, string arrayName, int index, object value)
        {
            SetError();

            Array ar = baseObject.GetType().InvokeMember(arrayName,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.GetField,
                null, baseObject, null) as Array;

            // *** Null assignments are not allowed because we may have to create the array
            // *** and a type is required for that. If necessary create an empty instance
            if (ar == null)
                return false;

            ar.SetValue(value, index);

            return true;
        }


        /// <summary>
        /// Removes an item from a .NET array with indirection to work around VFP's
        /// inability to manipulate .NET array elements.
        /// </summary>
        /// <param name="baseObject">The arrays parent object</param>
        /// <param name="arrayObject">The array's name as a string</param>
        /// <param name="Index">The index to of the item to delete. NOTE: 1 based!</param>
        /// <returns></returns>
        public bool RemoveArrayItem(object baseObject, string arrayObject, int Index)
        {
            SetError();

            // *** Assume 1 based arrays are passed in
            Index--;

            // *** USe Reflection to get a reference to the array Property
            Array ar = baseObject.GetType().InvokeMember(arrayObject,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.GetField,
                null, baseObject, null) as Array;

            // *** This may be ambigous - could mean no property or array exists and is null
            if (ar == null)
                return false;

            int arSize = ar.GetLength(0);

            if (arSize < 1)
                return false;

            Type arType = ar.GetValue(0).GetType();

            Array NewArray = Array.CreateInstance(arType, arSize - 1);

            // *** Manually copy the array
            int NewCount = 0;
            for (int i = 0; i < arSize; i++)
            {
                if (i != Index)
                {
                    NewArray.SetValue(ar.GetValue(i), NewCount);
                    NewCount++;
                }
            }

            baseObject.GetType().InvokeMember(arrayObject,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty | BindingFlags.SetField,
                null, baseObject, new object[1] {NewArray});

            return true;
        }

        /// <summary>
        /// Returns an XML string from a .NET DataSet
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="includeSchema"></param>
        /// <returns></returns>
        public string DataSetToXmlString(DataSet ds, bool includeSchema)
        {
            SetError();

            MemoryStream ms = new MemoryStream(4096);
            ds.WriteXml(ms, includeSchema ? XmlWriteMode.WriteSchema : XmlWriteMode.IgnoreSchema);
            string xml = Encoding.UTF8.GetString(ms.ToArray());
            return xml;
        }

        /// <summary>
        /// Converts an Xml String created from a FoxPro Xml Adapter or CursorToXml
        /// into a DataSet if possible.
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        public DataSet XmlStringToDataSet(string xml)
        {
            SetError();

            if (string.IsNullOrEmpty(xml))
                return null;

            try
            {
                using (MemoryStream ms = new MemoryStream(4096))
                {
                    byte[] xmlBuffer = Encoding.UTF8.GetBytes(xml);
                    ms.Write(xmlBuffer, 0, xmlBuffer.Length);
                    ms.Position = 0;

                    DataSet ds = new DataSet();
                    ds.ReadXml(ms);
                    return ds;
                }
            }
            catch (Exception ex)
            {
                SetError(ex.Message);
                return null;
            }
        }

        #endregion


        #region TypeConversionRoutines

        public static object FixupParameter(object val)
        {
            if (val == null || val is DBNull)
                return null;
            
            Type type = val.GetType();

            // *** Fix up binary SafeArrays into byte[]
            if (type.Name == "Byte[*]")
                return ConvertObjectToByteArray(val);

            // if we're dealing with ComValue parameter/value
            // just use it's Value property
            if (type == typeof(ComValue))
                return ((ComValue) val).Value;
            if (type == typeof(ComGuid))
                return ((ComGuid) val).Guid;
            if (type == typeof(ComArray))
                return ((ComArray) val).Instance;

            return val;
        }

        /// <summary>
        /// Fixes up a return value to return to FoxPro 
        /// based on its type. Fixes up some values to
        /// be type safe for FoxPro and others are returned
        /// as wrappers (ComArray, ComGuid)
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static object FixupReturnValue(object val)
        {
            if (val == null)
                return null;

            // special types to ignore and pass through
            if (val is string)
                return val;

            if(val is int || 
                val is bool || 
                val is DateTime)
                return val;

            if (val is ComValue || 
                val is Enum ||
                val is byte[])
                return val;

            // special handled types
            if (val is char)
                return val.ToString();

            if (val is long || val is Int64)
                return Convert.ToDecimal(val);

            if (val is Guid)
            {
                ComValue guidValue = new ComValue();
                guidValue.Value = (Guid) val;
                return guidValue;
            }


            // *** Need to figure out a solution for value types
            Type type = val.GetType();
            
            
            // FoxPro can't deal with DBNull as it's a value type
            if (type == typeof(DBNull))
            {
                val = null;
            }
            else if (type.IsArray)
            {
                ComArray comArray = new ComArray();
                comArray.Instance = val as Array;
                return comArray;
            }
            else if (type.IsGenericType && val is IList )
            {
                var enumerable = val as IEnumerable;
                ComArray comArray = new ComArray();
                comArray.FromEnumerable(enumerable);
                return comArray;
            }
            else if (type == typeof(Task) || type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Task<>))
            {
                // Minimize impact of deadlock when calling .Result or .Wait()
                var t = val as Task;                
                t.ConfigureAwait(false);
                return t;
            }
            else if (type.IsValueType && !type.IsPrimitive)
            {
                var comValue = new ComValue();
                comValue.Value = val;
                return comValue;
            }

            return val;
        }

        /// <summary>
        /// Converts an object to a byte array
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static byte[] ConvertObjectToByteArray(object val)
        {
            Array ct = (Array) val;
            Byte[] content = new byte[ct.Length];
            ct.CopyTo(content, 0);
            return content;
        }

        /// <summary>
        /// Helper routine that automatically assigns default names to certain
        /// 'common' system assemblies so that we don't have to provide a full path
        /// 
        /// NOTE: 
        /// All names are for .NET 2.0 Runtime at the moment
        /// </summary>
        /// <param name="AssemblyName"></param>
        protected string FixupAssemblyName(string AssemblyName)
        {
            string lowerAssemblyName = AssemblyName.ToLower();

            if (Environment.Version.Major == 2)
            {
                if (lowerAssemblyName == "system")
                    AssemblyName = "System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "mscorlib")
                    AssemblyName = "mscorlib, Version=2.1.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "system.windows.forms")
                    AssemblyName =
                        "System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "system.xml")
                    AssemblyName = "System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "system.drawing")
                    AssemblyName = "System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
                else if (lowerAssemblyName == "system.data")
                    AssemblyName = "System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "system.web")
                    AssemblyName = "System.Web, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
            }
            else if (Environment.Version.Major == 4)
            {
                if (lowerAssemblyName == "system")
                    AssemblyName = "System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "mscorlib")
                    AssemblyName = "mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "system.windows.forms")
                    AssemblyName =
                        "System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "system.xml")
                    AssemblyName = "System.Xml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "system.drawing")
                    AssemblyName = "System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
                else if (lowerAssemblyName == "system.data")
                    AssemblyName = "System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "system.web")
                    AssemblyName = "System.Web, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
                else if (lowerAssemblyName == "system.core")
                    AssemblyName = "System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "microsoft.csharp")
                    AssemblyName =
                        "Microsoft.CSharp, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
                else if (lowerAssemblyName == "microsoft.visualbasic")
                    AssemblyName =
                        "Microsoft.VisualBasic, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
                else if (lowerAssemblyName == "system.servicemodel")
                    AssemblyName =
                        "System.ServiceModel, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
                else if (lowerAssemblyName == "system.runtime.serialization")
                    AssemblyName =
                        "System.Runtime.Serialization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
            }

            return AssemblyName;
        }

        /// <summary>
        /// Returns a type reference of a .NET type (or FoxPro or COM object
        /// which is pretty much useless as it only returns the COM wrapper type)
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public Type GetType(object value)
        {
            if (value == null)
                return null;

            object fixedValue = null;
            if (value is ComValue)
                fixedValue = ((ComValue) value).Value;
            else
                fixedValue = value;
            
            return fixedValue.GetType();
        }

        /// <summary>
        /// Helper routine that looks up a type name and tries to retrieve the
        /// full type reference in the actively executing assemblies.
        /// </summary>
        /// <param name="typeName"></param>
        /// <returns></returns>
        public Type GetTypeFromName(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
                return null;

            Type type = null;
            foreach (Assembly ass in AppDomain.CurrentDomain.GetAssemblies())
            {
                type = ass.GetType(typeName, false);
                if (type != null)
                    break;
            }

            return type;
        }

        #endregion

        #region Error Reporting

        protected void SetError()
        {
            SetError(string.Empty);
        }

        protected void SetError(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                LastException = null;
                Error = false;
                ErrorMessage = string.Empty;
                return;
            }

            Error = true;
            ErrorMessage = message;

            if (IsThrowOnErrorEnabled)
                throw new InvalidOperationException(message);
        }

        protected void SetError(Exception ex, bool checkInner)
        {
            if (ex == null)
            {
                Error = false;
                ErrorMessage = string.Empty;
                LastException = null;
            }

            Exception e = ex;
            if (checkInner)
                e = ex.GetBaseException();

            Error = true;
            ErrorMessage = e.Message;
            LastException = e;

            if (IsThrowOnErrorEnabled)
                throw e;
        }

        protected void SetError(Exception ex)
        {
            SetError(ex, false);
        }

        #endregion

        #region Version

        public string GetVersionInfo()
        {
            bool isDotnetCore = false;
            var rt = GetStaticProperty("System.Runtime.InteropServices.RuntimeInformation", "FrameworkDescription");
            if (rt != null && Environment.Version.Major < 4)
                isDotnetCore = true;

            string res = null;
            if(!isDotnetCore) { 
            	res = $@"wwDotnetBridge Version   : {Assembly.GetExecutingAssembly().GetName().Version}
wwDotnetBridge Location  : {GetType().Assembly.Location}
.NET Version (official)  : {Environment.Version}
.NET Version (simplified): {GetDotnetVersion()}
.NET Version (Release)   : {GetDotnetVersion(true)}
Windows Version          : {GetWindowsVersion(WindowsVersionModes.Full)}";
            }
            else
            {
                res =  $@"wwDotnetBridge Version  : {Assembly.GetExecutingAssembly().GetName().Version}
wwDotnetBridge Location : {GetType().Assembly.Location}
.NET Core Version       : {Environment.Version}
.NET Core Version (Full): {rt}
Windows Version         : {GetWindowsVersion(WindowsVersionModes.Full)}";
            }

            return res;
        }


        static string DotnetVersion = null;

        /// <summary> 
        /// Returns the .NET framework version installed on the machine
        /// as a string  of 4.x.y version
        /// </summary>
        /// <remarks>Minimum version supported is 4.0</remarks>
        /// <returns></returns>
        public static string GetDotnetVersion(bool getReleaseVersion = false)
        {
            if (!string.IsNullOrEmpty(DotnetVersion) && !getReleaseVersion)
                return DotnetVersion;

            dynamic value;
            TryGetRegistryKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\", "Release", out value);

            if (value == null)
            {
                if (getReleaseVersion)
                    return null;

                DotnetVersion = "4.0";
                return DotnetVersion;
            }

            if (getReleaseVersion)
                return value.ToString();

            int releaseKey = value;

            // https://msdn.microsoft.com/en-us/library/hh925568(v=vs.110).aspx
            // RegEdit paste: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full
            if (releaseKey >= 528040) 
                DotnetVersion = "4.8";
            else if (releaseKey >= 461808)
                DotnetVersion = "4.7.2";
            else if (releaseKey >= 461308)
                DotnetVersion = "4.7.1";
            else if (releaseKey >= 460798)
                DotnetVersion = "4.7";
            else if (releaseKey >= 394802)
                DotnetVersion = "4.6.2";
            else if (releaseKey >= 394254)
                DotnetVersion = "4.6.1";
            else if (releaseKey >= 393295)
                DotnetVersion = "4.6";
            else if ((releaseKey >= 379893))
                DotnetVersion = "4.5.2";
            else if ((releaseKey >= 378675))
                DotnetVersion = "4.5.1";
            else if ((releaseKey >= 378389))
                DotnetVersion = "4.5";

            // This line should never execute. A non-null release key should mean 
            // that 4.5 or later is installed. 
            else
                DotnetVersion = "4.0";

            return DotnetVersion;
        }

        

        /// <summary>
        /// Returns Windows Version as a string
        /// 0 - Full: 10.0.17134.0 - Release: 1803
        /// 1 - Version only: 10.0.17134.0
        /// 2 - Release only: 1803
        /// </summary>
        /// <param name="mode"></param>
        /// <returns></returns>
        public string GetWindowsVersion(WindowsVersionModes mode = WindowsVersionModes.Full)
        {
            if (mode == WindowsVersionModes.VersionOnly)            
                return Environment.OSVersion.ToString();
            
            dynamic releaseId = null;
            TryGetRegistryKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
                "ReleaseId", out releaseId);

            if (mode == WindowsVersionModes.ReleaseOnly)
                return releaseId.ToString();
            
            return $"{Environment.OSVersion.Version}{(releaseId != null ? " - Release: " + releaseId : null)}";
        }

        internal static bool TryGetRegistryKey(string path, string key, out dynamic value, bool UseCurrentUser = false)
        {
            value = null;
            try
            {
                RegistryKey rk;
                if (UseCurrentUser)
                    rk = Registry.CurrentUser.OpenSubKey(path);
                else
                    rk = Registry.LocalMachine.OpenSubKey(path);

                if (rk == null) return false;
                value = rk.GetValue(key);
                return value != null;
            }
            catch
            {
                return false;
            }
        }


        #endregion


#if WestwindProduct

        // These are included only in the commercial version of Web Connection
        // the open source version omits these

        #region JsonXmlConversions

        /// <summary>
        /// Returns a JSON string from a .NET object. Note: doesn't
        /// work with FoxPro COM objects - only Interop .NET objects
        /// </summary>
        /// <param name="value"></param>
        /// <param name="formatted"></param>
        /// <returns></returns>
        public string ToJson(object value, bool formatted)
        {
            try
            {
                return JsonConvert.SerializeObject(
                    value,
                    formatted ? Formatting.Indented : Formatting.None);
            }
            catch (Exception ex)
            {
                SetError(ex, true);
                return null;
            }
        }

        /// <summary>
        /// Deserializes a JSON object
        /// </summary>
        /// <param name="json"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public object FromJson(string json, Type type)
        {
            try
            {
                return JsonConvert.DeserializeObject(json, type);
            }
            catch (Exception ex)
            {
                SetError(ex, true);
                return null;
            }
        }


        /// <summary>
        /// Returns an XML object from a .NET object. Note doesn't
        /// work with FoxPro COM object - only Interop .NET objects
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public string ToXml(object value)
        {
            string xmlResultString = null;

            try
            {
                SerializationUtils.SerializeObject(value, out xmlResultString, true);
            }
            catch (Exception ex)
            {
                SetError(ex, true);
                return null;
            }

            return xmlResultString;
        }


        /// <summary>
        /// Deserializes an object from an XML string that was 
        /// generated in format the same as generated from ToXml()
        /// </summary>
        /// <param name="xml"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public object FromXml(string xml, Type type)
        {
            try
            {
                return SerializationUtils.DeSerializeObject(xml, type);
            }
            catch (Exception ex)
            {
                SetError(ex, true);
                return null;
            }
        }

        /// <summary>
        /// Turns a date string into a JSON date string
        /// </summary>
        /// <param name="time">Time to JSON encode</param>
        /// <param name="isUtc">if false time is adjusted to UTC before serializing to remove timezone info</param>
        /// <remarks>
        /// There are problems with dates from a FoxPro table passed over COM
        /// into this function due to floating point rounding. Works fine with
        /// explicitly defined dates.
        /// </remarks>
        /// <returns></returns>
        public string ToJsonUtcDate(DateTime time, bool isUtc)
        {
            // fix rounding errors
            int second = time.Second;
            int minute = time.Minute;
            int hour = time.Hour;
            int millisecond = 0;
            if (time.Millisecond > 500)
            {
                second = time.Second + 1;
                if (second > 59)
                {
                    minute = minute + 1;
                    second = 0;
                }

                if (minute > 59)
                {
                    hour = hour + 1;
                    minute = 0;
                }

                if (hour > 23)
                {
                    hour = 23;
                    minute = 59;
                    second = 59;
                    millisecond = 999;
                }
            }

            // we need to fix the date because COM mucks up the milliseconds at times
            var dt = new DateTime(time.Year, time.Month, time.Day, hour, minute, second, millisecond);

            if (!isUtc)
                dt = dt.ToUniversalTime();
            var json = JsonConvert.SerializeObject(dt, wwJsonSerializer.jsonDateSettings);
            return json;
        }

        /// <summary>
        /// Returns the timezone offset of the local timezone
        /// to UTC in minutes including Daylight savings for
        /// the specific date.6
        /// </summary>
        /// <param name="localDate"></param>
        /// <returns></returns>
        public int GetLocalDateTimeOffset(DateTime localTime)
        {
            return Convert.ToInt32(new DateTimeOffset(localTime).Offset.TotalMinutes);
        }

        #endregion

#endif
    }


    public enum WindowsVersionModes
    {
        Full = 0,
        VersionOnly = 1,
        ReleaseOnly = 2
    }

}
