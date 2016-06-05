using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

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
            Thread.Sleep(2000);
            return "Hello world " + name;
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

        public string GetValues(string inputString, decimal inputDecimal)
        {
            return inputString + " " + inputDecimal.ToString();
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

        public char PassChar(char value)
        {
            return value;
        }

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

        public TestCustomer()
        {
            Address = new TestAddress();
        }
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
