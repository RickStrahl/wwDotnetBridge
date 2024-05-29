using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Westwind.WebConnection;

namespace wwDotnetBridge.Test
{
    [TestClass]
    public class ComCollectionsTests
    {

        public static string OutputLocation = Path.Combine(Path.GetDirectoryName(typeof(ComCollectionsTests).Assembly.Location));

        
        [TestMethod]
        public void CreateListTest()
        {
            var collection = new ComArray();
            var person = new Person();

            var typename = typeof(Person).FullName;
            Assert.IsTrue(collection.CreateList( typename), "Failed to create List<T>.");

            Assert.IsTrue(collection.AddItem(new Person { Name = "Rick", Company = "Westwind" }),
                "AddItem to list failed");

            Assert.AreEqual(collection.Count, 1);
        }

        [TestMethod]
        public void GetListIndexTest()
        {
            var list = new List<string>(new[] { "Test", "Test2" });

            var collection = new ComArray();
            collection.Instance = list;


            Assert.AreEqual(collection.Item(0) as string, "Test");
        }

        [TestMethod]
        public void AddListItemTest()
        {
            var list = new List<string>(new[] { "Test", "Test2" });

            var collection = new ComArray();
            collection.Instance = list;
            collection.Add("Test3");


            Assert.AreEqual(collection.Item(2) as string, "Test3");
        }

        [TestMethod]
        public void CreateListItemTest()
        {
            var list = new List<Person>(
                new[]
                {
                    new Person() { Name = "Rick", Company = "West Wind", Accessed = 10},
                    new Person() { Name = "Bill", Company = "Bonkers", Accessed = 12},
                });

            var collection = new ComArray();
            collection.Instance = list;

            var perrson = collection.CreateItem() as Person;

            Assert.IsNotNull(perrson);
        }


        [TestMethod]
        public void GetArrayIndexTest()
        {
            var list = new string[] { "Test", "Test2" };

            var collection = new ComArray();
            collection.Instance = list;

            Assert.AreEqual(collection.Item(0) as string, "Test");
        }

        [TestMethod]
        public void CreateArrayTest()
        {

            var collection = new ComArray();
            Assert.IsTrue(collection.CreateArray("System.String", 2));
            Assert.IsTrue(collection.Count == 2, "Collection should have two array items");
        }


        [TestMethod]
        public void CreateEmptyArrayTest()
        {
            var collection = new ComArray();
            Assert.IsTrue(collection.CreateArray("System.String", 0));
            Assert.IsNotNull(collection.Instance);
            Assert.IsTrue(collection.Count == 0, "Collection should have no array items");

            Assert.IsTrue(collection.Add("Test1"));
            Assert.IsTrue(collection.Count == 1, "Collection should have one array items");
        }

        [TestMethod]
        public void CreateArrayItemTest()
        {
            var list = new[]
                {
                    new Person() { Name = "Rick", Company = "West Wind", Accessed = 10},
                    new Person() { Name = "Bill", Company = "Bonkers", Accessed = 12},
                };

            var collection = new ComArray();
            collection.Instance = list;

            var person = collection.CreateItem() as Person;

            Assert.IsNotNull(person);
        }



        [TestMethod]
        public void AddArrayItemTest()
        {
            var list = new string[] { "Test", "Test2" };

            var collection = new ComArray();
            collection.Instance = list;

            collection.AddItem("Test3");

            Assert.AreEqual(collection.Item(2) as string, "Test3");
        }


        [TestMethod]
        public void CreateDictionaryTest()
        {
            var collection = new ComArray();
            var person = new Person();

            var typename = typeof(Person).FullName;
            Assert.IsTrue(collection.CreateDictionary("System.String", typename ), "Failed to create dictionary.");

            Assert.IsTrue(collection.AddDictionaryItem("Item1", new Person { Name = "Rick", Company = "Westwind" }),
                "AddDicitionaryItem failed");

            Assert.AreEqual(collection.Count, 1);

        }

        [TestMethod]
        public void GetDictionaryKeyTest()
        {
            var list = new Dictionary<string, int>();
            list.Add("Item1", 1);
            list.Add("Item2", 2);

            var collection = new ComArray();
            collection.Instance = list;

            var resultObj = collection.Item("Item1");

            Assert.AreEqual( resultObj, 1);
        }

        [TestMethod]
        public void CreateDictionaryItemTest()
        {
            var list = new Dictionary<string, Person>
            {
                {
                    "Item1",
                    new Person() { Name = "Rick", Company = "West Wind", Accessed = 10}
                },
                {
                    "Item2",
                    new Person()
                    {
                        Name = "Bill", Company = "Bonkers", Accessed = 12
                    }
                }
            };

            var collection = new ComArray();
            collection.Instance = list;

            var person = collection.CreateItem() as Person;

            Assert.IsNotNull(person);
        }

        [TestMethod]
        public void AddDictionaryKeyTest()
        {
            var list = new Dictionary<string,int>();
            list.Add("Item1", 1);
            list.Add("Item2", 2);


            var collection = new ComArray();
            collection.Instance = list;
            collection.AddDictionaryItem("Item3", 3);


            //collection.Item("Item3");

            Assert.AreEqual(collection.Item("Item3"), 3);
        }


        [TestMethod]
        public void GetHashSetKeyTest()
        {
            var list = new HashSet<string>();
            list.Add("Item1");
            list.Add("Item2");

            var collection = new ComArray();
            collection.Instance = list;


            Assert.AreEqual(collection.Item("Item1"), "Item1");
        }

        [TestMethod]
        public void CreateHashSetItemTest()
        {
            var list = new HashSet<Person>
            {
                {
                    new Person() { Name = "Rick", Company = "West Wind", Accessed = 10}
                },
                {
                    new Person()
                    {
                        Name = "Bill", Company = "Bonkers", Accessed = 12
                    }
                }
            };

            var collection = new ComArray();
            collection.Instance = list;

            var person = collection.CreateItem() as Person;

            Assert.IsNotNull(person);
        }

        [TestMethod]
        public void AddHashSetItemTest()
        {
            var list = new HashSet<string>();
            list.Add("Item1");
            list.Add("Item2");

            var collection = new ComArray();
            collection.Instance = list;

            collection.AddItem("Item3");

            Assert.AreEqual(collection.Item("Item3"), "Item3");
        }


    }

    public class Person
    {
        public string Name { get; set; }

        public string Company { get; set; }

        public int Accessed { get; set; }
    }

}