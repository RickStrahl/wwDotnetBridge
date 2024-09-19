using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Collections;
using System.Reflection;

namespace Westwind.WebConnection
{
    /// <summary>
    /// COM Wrapper for Arrays, IList and IDictionary collections that are not 
    /// directly or easily accessible in FoxPro. Provides members like `Count`, 
    /// `Item()` for index or key retrieval, `AddItem()`, `AddDictionaryItem()`, 
    /// `RemoveItem()`, `Clear()` to provide common collection operations more 
    /// naturally in FoxPro abstracted for the different types of collections 
    /// available in .NET (Arrays, Lists, Hash tables, Dictionaries).
    /// 
    /// Internally stores an instance of the collection so the collection is never 
    /// actually passed to FoxPro. Instead this class acts as a proxy to the 
    /// collection which is stored in the `.Instance` property.
    /// 
    /// This instance allows Visual FoxPro to manipulate the collection by directly
    ///  accessing array, list or dictionary values by index or key (depending on 
    /// list or collection), adding, removing and clearing items, creating array, 
    /// list and dictionary instances and assigning and picking up these  values 
    /// from existing objects.
    /// 
    /// This class is also internally used to capture collections that are returned
    ///  from method calls or Properties via `InvokeMethod()` and `GetProperty()` 
    /// and can be used to pass collections to .NET which otherwise would not be 
    /// possible over COM.
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
        public object Instance  {get; set; }

        /// <summary>
        /// Returns the length of the .NET array contained in Instance
        /// </summary>
        public int Count
        {
            get
            {
                if (Instance == null)
                    return 0;

                // these two should capture all list/array/collection scenarios
                if (Instance is ICollection)
                    return ((ICollection) Instance).Count;
                if (Instance is Array)
                    return ((Array)Instance).Length;

                // Try with Reflection - could fail with Exception
                return (int) ReflectionUtils.GetProperty(Instance, "Count");
            }

        }
        
        /// <summary>
        /// Default constructor
        /// </summary>
        public ComArray()
        {
        }

        /// <summary>
        /// Creates a new COM Array instance from an existing collection instance
        /// </summary>
        /// <param name="instance"></param>
        public ComArray(object instance)
        {
            Instance = instance;

            if (Instance == null)
                throw new ArgumentException("This instance is not an Array");

        }

        /// <summary>
        /// Returns an item by indexOrKey from an IList collection. This works on:
        /// 
        /// * Array   (int)
        /// * Array List  (int)
        /// * IList       (int)
        /// * ICollection  (any type)
        /// * HashSet      (any type)
        /// * IDictionary  (key type) (or int for index)
        /// 
        /// You can specify either an integer for list types or any type for 
        /// collections. int keys are 0 based.
        /// <seealso>Class ComArray</seealso>
        /// </summary>
        /// <returns>matching item or null</returns>
        /// <example>
        /// ```foxpro
        /// *** Access generic list through indirect access through Proxy and get 
        /// ComArray
        /// loList = loBridge.InvokeMethod(loNet,&quot;GetGenericList&quot;)
        /// ? loBridge.ToString(loList)   &&amp; System.Collections.Generic.List...
        /// ? loList.Count &&amp; 2
        /// 
        /// *** Grab an item by index
        /// loCust =  loList.Item(0)
        /// ? loCust.Company
        /// 
        /// *** Iterate the list
        /// FOR lnX = 0 TO loList.Count -1
        ///    loItem = lolist.Item(lnX)
        ///    ? loItem.Company
        /// ENDFOR
        /// ```
        /// 
        /// ```foxpro
        /// *** Retrieve a generic dictionary
        /// loList = loBridge.InvokeMethod(loNet,&quot;GetDictionary&quot;)
        /// ? loList.Count  &&amp;  2
        /// 
        /// *** Return Item by Key
        /// loCust =  loList.Item(&quot;Item1&quot;)   &&amp; Retrieve item by Key
        /// ? loCust.Company
        /// 
        /// *** This works as long as the key type is not int
        /// loCust =  loList.Item(0)   &&amp; Retrieve item by Index
        /// ? loCust.Company
        /// 
        /// *** This allows iterating a dictionary
        /// FOR lnX = 0 TO loList.Count -1
        ///    loItem = lolist.Item(lnX)
        ///    ? loItem.Company
        /// ENDFOR
        /// ```
        /// 
        /// </example>
        /// <remarks>
        /// Dictionaries can use 	an `int` value to return a value out of the 
        /// collection by its index **if the key type is not of `System.Int`. This is a
        ///  special case and allows you to iterate a dictionary which otherwise would 
        /// not be possible.
        /// </remarks>
        public object Item(object indexOrKey)
        {
            if (Instance == null)
                return null;

            if (Instance is IList)
            {
                if (!(indexOrKey is int))
                    throw new ArgumentException("List requires an integer key for lookups");

                var list = (IList)Instance;
                try
                {
                    object val = list[(int) indexOrKey];
                    if (val == null)
                        return null;

                    val = wwDotNetBridge.FixupReturnValue(val);
                    return val;
                }
                catch
                { }
            }

            if (Instance is IDictionary)
            {
                var list = (IDictionary)Instance;
                object val = null;
                try
                {
                   val = list[indexOrKey];
                }
                catch
                {
                }

				// Check for int key if the generic type is NOT int and return based on index
                if (val == null && indexOrKey is int && Instance.GetType().GenericTypeArguments[0] != typeof(int))
                {
                    int idx = (int)indexOrKey;
                    int counter = 0;
                    foreach (var item in list.Values)
                    {
                        if (idx == counter)
                            return item;
                        counter++;
                    }
                }

                if (val == null) return null;

                return wwDotNetBridge.FixupReturnValue(val);                
            }

            if (Instance is IEnumerable)
            {
                var list = Instance as IEnumerable;
                if (list == null) 
                    return null;

                foreach (var item in list)
                {
                    if (item == indexOrKey)
                        return wwDotNetBridge.FixupReturnValue(item);
                }

                // if an int was passed do a lookup to allow iteration
                if (indexOrKey is int)
                {
                    int idx = (int)indexOrKey;
                    int counter = 0;
                    foreach (var item in list)
                    {
                        if (idx == counter)
                            return wwDotNetBridge.FixupReturnValue(item);
                        counter++;
                    }
                }
                
                return null;
            }

            return null;
        }


        /// <summary>
        /// Returns the indexed item without any type conversion fixup
        /// </summary>
        /// <param name="indexOrKey"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public object ItemRaw(object indexOrKey)
        {
            if (Instance == null)
                return null;

            if (Instance is IList)
            {
                if (!(indexOrKey is int))
                    throw new ArgumentException("List requires an integer key for lookups");

                var list = (IList)Instance;
                try
                {
                    object val = list[(int)indexOrKey];
                    if (val == null)
                        return null;
                    return val;
                }
                catch
                { }
            }

            if (Instance is IDictionary)
            {
                var list = (IDictionary)Instance;
                object val = null;
                try
                {
                    val = list[indexOrKey];
                }
                catch
                {
                }

                // Check for int key if the generic type is NOT int and return based on index
                if (val == null && indexOrKey is int && Instance.GetType().GenericTypeArguments[0] != typeof(int))
                {
                    int idx = (int)indexOrKey;
                    int counter = 0;
                    foreach (var item in list.Values)
                    {
                        if (idx == counter)
                            return item;
                        counter++;
                    }
                }

                if (val == null) return null;

                return val;
            }

            if (Instance is IEnumerable)
            {
                var list = Instance as IEnumerable;
                if (list == null) 
                    return null;

                foreach (var item in list)
                {
                    if (item == indexOrKey)
                        return item;
                }

                // if an int was passed do a lookup to allow iteration
                if (indexOrKey is int)
                {
                    int idx = (int)indexOrKey;
                    int counter = 0;
                    foreach (var item in list)
                    {
                        if (idx == counter)
                            return item;
                        counter++;
                    }
                }
                
                return null;
            }

            return null;
        }


        /// <summary>
        /// Returns the type name of the array instance
        /// </summary>
        /// <returns></returns>
        public string GetInstanceTypeName()
        {
            if (Instance == null)
                return string.Empty;

            return Instance?.GetType().FullName ?? string.Empty;
        }

        /// <summary>
        /// Returns the type name of each element in the array instance
        /// </summary>
        /// <returns></returns>
        public string GetItemTypeName()
        {
            if (Instance == null)
                return string.Empty;

            if (Instance is Array)
            {
                Array ar = Instance as Array;
                if (ar.Length < 1)
                    return string.Empty;
                return Instance.GetType().GetElementType()?.FullName ?? string.Empty;
            }

            if (Instance is IList)
            {
                IList i  = Instance as IList;
                if (i.Count < 1)
                    return string.Empty;
                return Instance.GetType().GetGenericArguments()[0].FullName;
            }

            if (Instance is IDictionary)
            {
                IDictionary i = Instance as IDictionary;
                if (i.Count < 1)
                    return string.Empty;
                return Instance.GetType().GetGenericArguments()[1].FullName;
            }

            return string.Empty;
        }
       

        /// <summary>
        /// Adds an item to the internal list, array or single item collection.
        /// 
        /// The list or collection has to exist or an exception is thrown.
        /// </summary>
        /// <param name="item">an instance of the item to add.</param>
        /// <returns>true or false</returns>
        public bool AddItem(object item)
        {
            item = wwDotNetBridge.FixupParameter(item);

            if (Instance is Array)
            {
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
                    if (!CreateArray(itemType.FullName,0))
                        return false;

                    ar = Instance as Array;
                }

                int size = ar.GetLength(0);
                Array copiedArray = Array.CreateInstance(ar.GetType().GetElementType(), size + 1);
                ar.CopyTo(copiedArray, 0);

                ar = copiedArray;
                ar.SetValue(item, size);

                Instance = ar;

                return true;
            }
            if (Instance is IList)
            {
                var list = Instance as IList;
                list.Add(item);
                return true;
            }

            try
            {
                // force it - for special collections or HashSet<T>
                ReflectionUtils.CallMethod(Instance, "Add", item);
                return true;
            }
            catch
            { }

            return false;
        }

        /// <summary>
        /// Adds an item to the internal list or collection/dictionary.
        /// 
        /// The list or collection has to exist or an exception is thrown.
        /// </summary>
        /// <param name="item">an instance of the item to add.</param>
        /// <returns>true or false</returns>
        public bool Add(object item)
        {
            return AddItem(item);
        }

        /// <summary>
        /// Adds an item to the dictionary by key and value.
        /// 
        /// **The list or collection has to exist or an exception is thrown.**
        /// </summary>
        /// <param name="key">Dictionary key used to look up this item.</param>
        /// <param name="value">an instance of the item to add.</param>
        /// <returns>true or false</returns>
        public bool AddDictionaryItem(object key, object value)
        {
            if (!(Instance is IDictionary)) return false;

            var list = (IDictionary)Instance;
            list.Add(key, value);
            return true;
        }

        /// <summary>
        /// Updates a value to an existing collection or list item
        /// by int index for lists, or a key or value for dictionaries/collections
        /// </summary>
        /// <param name="indexOrKey">int 0 based index or any type for key on collections/dictionaries</param>
        /// <param name="value"></param>
        /// <returns>true or false</returns>
        public bool SetItem(object indexOrKey, object value)
        {
            value = wwDotNetBridge.FixupParameter(value);

            if (Instance == null)
                return false;

            if (Instance is Array)
            {
                if (!(indexOrKey is int))
                    throw new ArgumentException("Array indexOrKey has to be integer");

                var list = Instance as Array;
                list.SetValue(value, (int)indexOrKey);
                return true;
            }
            if (Instance is IList)
            {
                if (!(indexOrKey is int))
                    throw new ArgumentException("List indexOrKey has to be integer");

                var list = Instance as IList;
                list[(int) indexOrKey] = value;
                return true;
            }
            if (Instance is IDictionary)
            {
                var list = Instance as IDictionary;
                list[indexOrKey] = value;
                return true;
            }
            

            return false;
        }

        /// <summary>
        /// Removes an item from the collection.
        ///
        /// Depending on the item is a List style item (Array, List, HashSet, simple Collections)
        /// or a Key Value type collection (Dictionary or custom key value collections like StringCollection),
        /// you specify an index (lists), value (Hash tables/collections), or a key (Dictionary).
        /// </summary>
        /// <param name="indexOrKey">0 based list index or value for a Hash table, or  a key to a collection/dictionary</param>
        /// <returns>true or false</returns>
        public bool RemoveItem(object indexOrKey)
        {
            // Remove by indexOrKey
            if (Instance is IList)
            {
                var list = Instance as IList;
                list.Remove(indexOrKey);
                return true;
            }
            // Remove by key
            if (Instance is IDictionary)
            {
                var list = Instance as IDictionary;
                list.Remove(indexOrKey);
                return true;
            }

            if (Instance is IDictionary)
            {
                var list = Instance as IDictionary;
                list.Remove(indexOrKey);
                return true;
            }
            if (Instance is Array)
            {
                if (!(indexOrKey is int))
                    throw new ArgumentException("Array indexers to remove have to be of type int");

                int idx = (int) indexOrKey;

                // *** USe Reflection to get a reference to the array Property
                Array ar = Instance as Array;
                int arSize = ar.GetLength(0);

                if (arSize < 1)
                    return false;

                Type arType = ar.GetType().GetElementType();

                Array newArray = Array.CreateInstance(arType, arSize - 1);

                // *** Manually copy the array
                int newArrayCount = 0;
                for (int i = 0; i < arSize; i++)
                {
                    if (i != idx)
                    {
                        newArray.SetValue(ar.GetValue(i), newArrayCount);
                        newArrayCount++;
                    }
                }

                Instance = newArray;
                return true;
            }

            ReflectionUtils.CallMethod(Instance, "Remove", indexOrKey);
            return true;
        }

        /// <summary>
        /// Removes an item from the collection.
        ///
        /// Depending on the item is a List style item (Array, List, HashSet, simple Collections)
        /// or a Key Value type collection (Dictionary or custom key value collections like StringCollection),
        /// you specify an index (lists), value (Hash tables/collections), or a key (Dictionary).
        /// </summary>
        /// <param name="indexOrKey">0 based list index or value for a Hash table, or  a key to a collection/dictionary</param>
        /// <returns>true or false</returns>
        public bool Remove(object indexOrKey)
        {
            return RemoveItem(indexOrKey);
        }

        /// <summary>
        /// Clears out the collection's content.
        /// </summary>
        /// <returns></returns>
        public bool Clear()
        {
            if (Instance is Array)
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

            if (Instance == null)
                return false;

            if (Instance is IList)
                ((IList)Instance).Clear();
            else if (Instance is IDictionary)
                ((IList)Instance).Clear();
            else
                ReflectionUtils.CallMethod(Instance, "Clear");
            
            return false;
        }

        /// <summary>
        /// Creates an instance of the collection's data member type without
        /// actually adding it to the array. This is useful to
        /// quickly create empty objects that can be populated in FoxPro
        /// without specifying the full typename each time resulting in
        /// cleaner code in FoxPro. 
        /// 
        /// **Assumes that the collection `Instance` is set or null is returned**
        /// 
        /// Arrays must have at least 1 item, while generic lists can
        /// deduce from the Generic Type parameters.
        /// </summary>
        /// <returns>Item created or null, or if no items or unable to create (or object</returns>
        public object CreateItem()
        {
            if (Instance == null)
                return null;

            var type = Instance.GetType();

            Type itemType = null;
            if (type.GenericTypeArguments.Length > 1)
            {
                itemType = type.GenericTypeArguments[1];
            }
            else if(type.GenericTypeArguments.Length == 1)
            {
                itemType = type.GenericTypeArguments[0];
            }
            else if (Instance is Array)
            {
                itemType = Instance.GetType().GetElementType();
            }
            else
            {
                itemType = typeof(object);
            }

            object item = ReflectionUtils.CreateInstanceFromType(itemType);
            return item;
        }

        /// <summary>
        /// Creates a new array instance with size number
        /// of items pre-set. Elements are unassigned but
        /// array is dimensioned. Use 0 for an empty array.
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
        /// Creates an empty List of T.
        /// </summary>
        /// <param name="listTypeName">.NET Type of the items in the List</param>
        /// <returns></returns>
        public bool CreateList(string listTypeName)
        {
            try
            {
                var type = ReflectionUtils.GetTypeFromName(listTypeName);
                var genericType = typeof(List<>).MakeGenericType(type);
                Instance = Activator.CreateInstance(genericType) as IList;
            }
            catch
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Creates an empty Dictionary  with specified key value types
        /// </summary>
        /// <param name="keyTypeName">.NET Typename of the Key</param>
        /// <param name="valueTypeName">.NET Typename of the collection Data items</param>
        /// <returns></returns>
        public bool CreateDictionary(string keyTypeName, string valueTypeName)
        {
            try
            {
                var valueType = ReflectionUtils.GetTypeFromName(valueTypeName);
                var keyType = ReflectionUtils.GetTypeFromName(keyTypeName);

                var genericType = typeof(Dictionary<,>).MakeGenericType(keyType, valueType);
                Instance = Activator.CreateInstance(genericType) as IDictionary;
            }
            catch
            {
                return false;
            }

            return true;
        }
        

        /// <summary>
        /// Assigns a .NET collect to this ComArray Instance. Has to be passed
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
            Instance = ReflectionUtils.GetPropertyEx(baseInstance, arrayPropertyName);
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
        /// Creates a copied array instance from an enumerable.
        ///
        /// If enumerables aren't bound to a concrete list/collection
        /// and retrieved one at a time, there's no way to return it to
        /// FoxPro unless it's materialized into some sort of list.
        ///
        /// This method makes a copy of the enumerable and allows you to
        /// iterate over it.
        /// <remarks>
        /// You can update existing items in the array, but you cannot
        /// add/remove/clear items and affect the underlying enumerable.
        /// If the type is truly Enumerated it won't support updating
        /// either, so this is not really a restriction imposed here.
        ///
        /// If the underlying enumerable is a concrete list or collection 
        /// type you should be able to use the relevant Add/Remove/Clear
        /// methods directly without requiring this method to make a copy.
        /// </remarks>
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
