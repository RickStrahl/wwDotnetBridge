#if DEBUG
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace Westwind.WebConnection
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class TypePassingTests
    {
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

        public int PassCustomerArray(TestCustomer[] customers)
        {
            if (customers != null)
                return customers.Length;

            return 0;
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

}
#endif