#if NETFULL

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Data;
using System.Threading.Tasks;
using System.Web.UI.WebControls;

namespace Westwind.WebConnection
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class TypePassingTests
    {
        public decimal? DecimalValue { get; set; }

        public UInt64   Uint64Value { get; set; }

        public TypePassingTests()
        {
            DecimalValue = 1.0M;
        }

        public string HelloWorld(string name)
        {
            return "Hello world " + name;
        }

        private string PrivateHelloWorld(string name)
        {
            return "Hello " + name;
        }

        public static string HelloWorldStatic(string name)
        {
            Thread.Sleep(3000);
            return "Hello static world " + name;
        }

        public Guid GetGuid()
        {
            return Guid.NewGuid();
        }

        public string SetGuid(Guid guid)
        {
            return guid.ToString();
        }

        public Session GetGuidOnObject()
        {
            var session = new Session() {Guid = Guid.NewGuid()};
            return session;
        }

        public class Session
        {
            public Guid Guid { get; set; }
        }




        public string GetValues(string inputString, decimal inputDecimal)
        {
            return inputString + " " + inputDecimal.ToString();
        }


        public DataSet GetDataSet()
        {
            var dataSet = new DataSet();

            var dataTable = new DataTable();
            dataTable.Columns.Add("name",typeof(string));
            dataTable.Columns.Add("company",typeof(string));
            dataTable.Columns.Add("entered",typeof(DateTime));

            var dataRow = dataTable.NewRow();

            dataRow["name"] = "rick";
            dataRow["company"] = "West Wind";
            dataRow["Entered"] = DateTime.Now;
            dataTable.Rows.Add(dataRow);

            dataRow = dataTable.NewRow();

            dataRow["name"] = "Markus";
            dataRow["company"] = "EPS Software";
            dataRow["Entered"] = DateTime.Now.AddDays(1);
            dataSet.Tables.Add(dataTable);

            return dataSet;
        }

        public Int16 PassInt16(Int16 val)
        {
            return val;
        }

        public byte[] PassBinary(byte[] binData)
        {
            return binData;
        }


        public byte PassByte(byte byteData)
        {
            return byteData;
        }

        public Single PassSingle(Single number)
        {
            return number + 10.2F;
        }

        public Decimal PassDecimal(decimal number)
        {
            return number + 11.2M;
        }

        public long PassLong(long number)
        {
            return number + 10;
        }

        public long[] PassLongArray(long[] numbers)
        {
            return numbers;
        }


        /// <summary>
        /// Should return as decimal
        /// </summary>
        /// <returns></returns>
        public long ReturnLong()
        {
            // 1000 trillion
            return 1_000_000_000_000_001; 
        }


        public float ReturnFloat()
        {
            return 1.43123F;
        }


        public Single ReturnSingle()
        {
            return 1.43123F;
        }

        public Double ReturnDouble()
        {
            return 1.4333333123123F;
        }

        public List<TestCustomer> GetGenericList()
        {
            var list = new List<TestCustomer>();

            list.Add(new TestCustomer() { Company = "West Wind"});
            list.Add(new TestCustomer() { Company = "East Wind" });

            return list;
        }

        public int SetGenericList(List<TestCustomer> list)
        {
            return list.Count;
        }

        public Dictionary<string, TestCustomer> GetDictionary()
        {
            var list = new Dictionary<string, TestCustomer>();
            list.Add("Item1", new TestCustomer() { Company = "West Wind" });
            list.Add("Item2",new TestCustomer() { Company = "East Wind" });

            return list;
        }

        public int SetDictionary(Dictionary<string,TestCustomer> dict)
        {
            return dict.Count;
        }

        public char PassChar(char value)
        {
            return value;
        }

        /// <summary>
        /// This should be returned as a ComValue object
        /// </summary>
        /// <param name="intVal"></param>
        /// <param name="stringVal"></param>
        /// <returns></returns>
        public StructValue ReturnStruct(int intVal, string stringVal)
        {
            return new StructValue
            {
                IntValue = intVal,
                StringValue = stringVal
            };
        }



        public static StructValue ReturnStructStatic(int intVal, string stringVal)
        {
            return new StructValue
            {
                IntValue = intVal,
                StringValue = stringVal
            };
        }

        public TestCustomer[] GetCustomerArray()
        {
            TestCustomer[] customers = new TestCustomer[2];

            TestCustomer cust = customers[1];
            cust.Name = "Rick Strahl";
            cust.Company = "West Wind";
            cust.Entered = DateTime.Now;
            cust.Address.City = "Paia";
            cust.Address.Street = "32 Kaiea";

            cust = customers[2];
            cust.Name = "Jimmy Row";
            cust.Company = "East Wind";
            cust.Entered = DateTime.Now;
            cust.Address.City = "Hood River";
            cust.Address.Street = "301 15th Street";

            return customers;
        }


        public void PassByReference(ref int intValue, ref string stringValue, ref decimal decimalValue)
        {
            intValue = intValue*2;
            decimalValue = decimalValue*2;
            stringValue += " Updated!";
        }

        public void PassArrayByReference(ref string[] stringValues)
        {
            var list = new List<string>(stringValues);
            list.AddRange(new string[] { "One", "More", "set", "of", "words" });
            stringValues = list.ToArray();
        }


        public static void PassByReferenceStatic(ref int intValue, ref string stringValue, ref decimal decimalValue)
        {
            intValue = intValue*2;
            decimalValue = decimalValue*2;
            stringValue += " Updated!";
        }



        public int PassCustomerArray(TestCustomer[] customers)
        {
            if (customers != null)
                return customers.Length;

            return 0;
        }

        /// <summary>
        /// </summary>
        /// <example>
        /// LOCAL loArray as Westwind.WebConnection.ComArray
        /// loArray = loBridge.Createarray("System.Windows.Forms.MessageBoxButtons")
        /// loComValue = loBridge.CreateComValue()
        /// loComValue.SetEnum("System.Windows.Forms.MessageBoxButtons.OK")
        /// loArray.AddItem( loComValue) 
        /// ? loBridge.InvokeMethod(loTest,"PassEnumArray",loArray)
        /// </example>
        /// <param name="buttons"></param>
        /// <returns></returns>
        public int PassEnumArray(MessageBoxButtons[] buttons)
        {
            if (buttons == null)
                return 0;

            return buttons.Length;
        }

        /// <summary>
        /// Pass in 15 parameters and return the value of the last
        /// </summary>
        /// <returns></returns>
        public int Pass24Parameters(int p1, int p2, int p3, int p4, int p5, int p6, int p7,
                                    int p8, int p9, int p10, int p11, int p12, int p13, int p14, int p15,
                                    int p16, int p17, int p18, int p19, int p20,
                                    int p21, int p22, int p23, int p24)
        {
            return p24;
        }

        public int PassUnlimitedParameters(params Int32[] numbers)
        {
            return numbers.Length;
        }


        public async Task<int> AddAsync(int num1, int num2)
        {
            await Task.Delay(1000);
            return num1 + num2;
        }

        public async Task<int> ThrowErrorAsync()
        {
            await Task.Delay(1000);

            throw new Exception("Throwing Copper - Asynchronously. Should cause an exception in async code");

        }

    }



    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class TestCustomer
    {
        public string Name { get; set; }
        public string Company { get; set; }
        public DateTime Entered { get; set; }
        public decimal Number { get; set; }
        public TestAddress Address { get; set; }

        public TestItem[] Items { get; set; } = 
        {
            new TestItem {Key = "Item1", Value = "Value 1"}, 
            new TestItem {Key = "Item2", Value = "Value 2"}
        };
        
        public TestCustomer()
        {
            Address = new TestAddress();
        }

        public override string ToString()
        {
            return $"{Name} {Company}";
        }
    }

    public class TestItem
    {
        public string Key { get; set; }
        public string Value {get; set; }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class TestAddress
    {
        public string Street { get; set; }
        public string City { get; set; }

        public TestAddress()
        {
            Street = string.Empty;
            City = string.Empty;
        }
    }

    public struct StructValue
    {
        public int IntValue { get; set; }
        public string StringValue { get; set; }
    }

}
#endif