using System;
using System.Runtime.InteropServices;

namespace Westwind.WebConnection
{
    /// <summary>
    /// .NET System.Guid values cannot be passed to VFP as they are non COM 
    /// exported value types. So any Guid values passed to and received from .NET 
    /// need to be passed around as a ComGuid instance. This class wraps an 
    /// internal Guid instance member and allows access to the GuidString property 
    /// via string.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Obsolete("Please replace ComGuid with ComValue")]
    [ProgId("Westwind.WebConnection.ComGuid")]
    public class ComGuid
    {
        /// <summary>
        /// The actual Guid instance that can be read by a .NET handler
        /// </summary>
        public Guid Guid
        {
            get { return _Guid; }
            set { _Guid = value; }
        }
        private Guid _Guid = Guid.Empty;


        /// <summary>
        /// String representation of the Guid. 
        /// </summary>
        public string GuidString
        {
            get
            {
                return Guid.ToString();
            }
            set
            {
                Guid = new Guid(value);
            }
        }

        /// <summary>
        /// Sets the Guid instance to a new Guid Value
        /// </summary>
        public void New()
        {
            Guid = Guid.NewGuid();
        }

        /// <summary>
        /// Sets the Guid to Guid.Empty
        /// </summary>
        public void Empty()
        {
            Guid = Guid.Empty;
        }

    }
}
