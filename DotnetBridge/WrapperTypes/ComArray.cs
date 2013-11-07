using System;
using System.Runtime.InteropServices;
using Westwind.Utilities;
using System.Collections.Generic;
using System.Collections;
using System.Reflection;

namespace Westwind.WebConnection
{
    /// <summary>
    /// COM Wrapper for an array that is assigned as variable.
    /// This instance allows Visual FoxPro to manipulate the array
    /// using the wwDotNetBridge Array functions that are require
    /// a parent object.
    /// 
    /// When passed to a method that requires an array the instance
    /// member is passed as the actual parameter.
    /// 
    /// Note: You should always use wwDotNetBridge.CreateInstance
    /// to create an instance of this array from Fox code otherwise
    /// there's no instance set.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("Westwind.WebConnection.ComArray")]
    public class ComArray
    {
        /// <summary>
        /// The actual array instance returned as an object.
        /// This instance is set and passed to and from .NET
        /// calls made with InvokeMethod and explicit property
        /// assignments with Set/GetProperty.
        /// </summary>
        public object Instance
        {
            get { return _Instance; }
            set { _Instance = value; }
        }
        private object _Instance = null;

        /// <summary>
        /// Returns the length of the .NET array contained in Instance
        /// </summary>
        public int Count
        {
            get
            {
                if (Instance == null)
                    return 0;

                Array inst = Instance as Array;
                if (inst == null)
                    return 0;

                return inst.Length;
            }

        }
		

        /// <summary>
        /// Default constructor
        /// </summary>
        public ComArray()
        {
        }

        /// <summary>
        /// Creates a new COM Array from an existing array instance
        /// </summary>
        /// <param name="instance"></param>
        public ComArray(object instance)
        {
            Instance = instance as Array;

            if (Instance == null)
                throw new ArgumentException("This instance is not an Array");

        }

        /// <summary>
        /// Creates a .NET array instance with 0 items on this ComArray instance
        /// </summary>
        /// <param name="arrayTypeName"></param>
        /// <returns></returns>
        public bool CreateEmptyArray(string arrayTypeName)
        {
            return CreateArray(arrayTypeName,0);
        }
        
        /// <summary>
        /// Deprecated: Don't use
        /// </summary>
        /// <param name="arrayTypeName"></param>
        /// <returns></returns>        
        public bool Create(string arrayTypeName)
        {
            return CreateArray(arrayTypeName,0);
        }

        /// <summary>
        /// Creates a new array instance with size number
        /// of items pre-set. Elements are unassigned but
        /// array is dimensioned.
        /// 
        /// Use SetItem() to assign values to each array element
        /// </summary>
        /// <param name="arrayTypeName">The type of the array's elements</param>
        /// <param name="size">Size of array to create</param>
        /// <returns></returns>
        public bool CreateArray(string arrayTypeName, int size)
        {
            Type type = null;
            type = ReflectionUtils.GetTypeFromName(arrayTypeName);

            if (type == null)
                return false;

            // *** Create instance and assign
            Array ar = Array.CreateInstance(type, size);

            if (ar == null)
                return false;

            Instance = ar;

            return true;            
        }

        /// <summary>
        /// Assigns a .NET array to this COM wrapper. Has to be passed
        /// as a base instance (ie. parent instance of the array) and
        /// the name of the array because once the array hits VFP code
        /// it's already been converted into a VFP array so only internal
        /// reflection will allow getting the actual reference into ComArray.
        /// </summary>
        /// <param name="baseInstance">Instance of the parent object of the array</param>
        /// <param name="arrayPropertyName">Name of the array property on the parent instance</param>
        /// <returns></returns>
        public bool AssignFrom(object baseInstance, string arrayPropertyName)
        {            
            Instance = ReflectionUtils.GetPropertyEx(baseInstance, arrayPropertyName) as Array;                       
            return true;
        }

        /// <summary>
        /// Assigns this ComArray's array instance to the specified property
        /// </summary>
        /// <param name="baseInstance"></param>
        /// <param name="arrayPropertyName"></param>
        /// <returns></returns>
        public bool AssignTo(object baseInstance, string arrayPropertyName)
        {
            ReflectionUtils.SetPropertyEx(baseInstance, arrayPropertyName, Instance);
            return true;
        }

        /// <summary>
        /// Creates an instance of the array's member type without
        /// actually adding it to the array. This is useful to
        /// more easily create members without having to specify
        /// the full type signature each time.
        /// 
        /// Assumes that the array exists already so that the 
        /// item type can be inferred. The type is inferred from
        /// the arrays instance using GetElementType().
        /// </summary>
        /// <returns></returns>
        public object CreateItem()
        {
            if (Instance == null)
                return null;

            Type itemType = Instance.GetType().GetElementType();

            object item = ReflectionUtils.CreateInstanceFromType(itemType);
            return item;
        }

        /// <summary>
        /// Creates an instance of the array's member type without
        /// actually adding it to the array. This is useful to
        /// more easily create members without having to specify
        /// the full type signature each time.
        /// 
        /// This version works of the actual elements in the array
        /// instance rather than using the 'official' element type.
        /// Looks at the first element in the array and uses its type.
        /// 
        /// Assumes that the array exists already so that the 
        /// item type can be inferred.
        /// </summary>
        /// <param name="forceElementType">If true looks at the first element and uses that as the type to create.
        /// Use this option if the actual element type is of type object when the array was automatically generated
        /// such as when FromEnumerable() was called.
        /// </param>
        /// <returns></returns>
        public object CreateItemExplicit()
        {
            if (Instance == null)
                return null;

            Type itemType = null;
            var arInst = Instance as Array;
            if (arInst != null && arInst.Length > 0)
               itemType = arInst.GetValue(0).GetType();            
            if (itemType == null)
               itemType = Instance.GetType().GetElementType();

            object item = ReflectionUtils.CreateInstanceFromType(itemType);
            return item;
        }

        /// <summary>
        /// Returns an item from the array.
        /// </summary>
        /// <param name="index">0 based array index to retrieve</param>
        /// <returns></returns>
        public object Item(int index)
        {
            if (Instance == null)
                return null;

            Array ar = Instance as Array;

            try
            {
                object val = ar.GetValue(index);
                if (val == null)
                    return null;

                wwDotNetBridge.FixupReturnValue(val);

                return val;
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Assigns a value to an array element that already exists.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool SetItem(int index, object value)
        {
            Array ar = Instance as Array;
            ar.SetValue(value, index);

            return true;
        }

        /// <summary>
        /// Adds an item to the internal array instance.
        /// 
        /// Array should exist before adding items.
        /// </summary>
        /// <param name="item">an instance of the item to add.</param>
        /// <returns></returns>
        public bool AddItem(object item)
        {
            item = wwDotNetBridge.FixupParameter(item);
            
            Array ar = Instance as Array;

            Type itemType = null;
            if (item != null)
                itemType = item.GetType();
            else if (ar != null)
                itemType = ar.GetType().GetElementType();
            if (itemType == null)
                return false;  

            // *** This may be ambigous - could mean no property or array exists and is null
            if (ar == null)
            {
                if (!CreateEmptyArray(itemType.FullName))
                    return false;
            }
                        
            int size = ar.GetLength(0);
            Array copiedArray = Array.CreateInstance(ar.GetType().GetElementType(), size + 1);
            ar.CopyTo(copiedArray, 0);

            ar = copiedArray;
            ar.SetValue(item, size);

            Instance = ar;

            return true;
        }

        /// <summary>
        /// Removes an item from the array.
        /// </summary>
        /// <param name="index">0 based index of item to remove</param>
        /// <returns></returns>
        public bool RemoveItem(int index)
        {
            // *** USe Reflection to get a reference to the array Property
            Array ar = Instance as Array; 
            
            // *** This may be ambigous - could mean no property or array exists and is null
            if (ar == null)
                return false;

            int arSize = ar.GetLength(0);

            if (arSize < 1)
                return false;            

            Type arType = ar.GetType().GetElementType();

            Array newArray = Array.CreateInstance(arType, arSize - 1);

            // *** Manually copy the array
            int newArrayCount = 0;
            for (int i = 0; i < arSize; i++)
            {
                if (i != index)
                {
                    newArray.SetValue(ar.GetValue(i), newArrayCount);
                    newArrayCount++;
                }
            }
            Instance = newArray;
            return true;
        }

        /// <summary>
        /// Clears out the array contents
        /// </summary>
        /// <returns></returns>
        public bool Clear()
        {
            Array ar = Instance as Array;
            if (ar == null)
                return true;  

            if (ar.GetLength(0) < 1)
                return true;

            var type = ar.GetValue(0).GetType();
            Instance = Array.CreateInstance(type, 0);


            return true;
        }

        /// <summary>
        /// Creates an instance from an enumerable
        /// </summary>
        /// <param name="items"></param>
        public void FromEnumerable(IEnumerable items)
        {
            ArrayList al = new ArrayList();
            foreach (var item in items)
                al.Add(item);

            if (al.Count < 1)
            {
                Instance = null;
                return;
            }

            Type elType = al[0].GetType();
            Array ar = Array.CreateInstance(elType,al.Count);
            for (int i = 0; i < al.Count; i++)
            {
            	ar.SetValue(al[i],i);
            }
            Instance = ar;
        }
    }


}
