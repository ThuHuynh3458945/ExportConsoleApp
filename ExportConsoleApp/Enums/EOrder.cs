using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportConsoleApp.Enums
{
    public enum EOrderStatusExtend
    {
        /// <summary>
        /// All Orders
        /// </summary>
        [Description("All")]
        All = 0,
        /// <summary>
        /// Error
        /// </summary>
        [Description("Error")]
        Error = 1,
        /// <summary>
        /// Alert Admin
        /// </summary>
        [Description("AlertAdmin")]
        AlertAdmin = 2,
        /// <summary>
        /// this order status tab will be do later
        /// </summary>
        ReprintRequests = 3,
        /// <summary>
        /// this order status tab will be do later
        /// </summary>
        [Description("OnHold")]
        OnHold = 4,
        /// <summary>
        /// Unassigned
        /// </summary>
        [Description("Unassigned")]
        Unassigned = 5,
        /// <summary>
        /// UnShipped
        /// </summary>
        [Description("Unshipped")]
        Unshipped = 6,
        /// <summary>
        /// In Shortage
        /// </summary>
        [Description("InShortage")]
        InShortage = 7,
        /// <summary>
        /// Not Dropped
        /// </summary>
        [Description("NotDropped")]
        NotDropped = 8,
        /// <summary>
        /// In Progress
        /// </summary>
        [Description("InProgress")]
        InProgress = 9,
        /// <summary>
        /// Packaged (ReadyToShip)
        /// </summary>
        [Description("Packaged")]
        Packaged = 10,
        /// <summary>
        /// Shipped
        /// </summary>
        [Description("Shipped")]
        Shipped = 11,
        /// <summary>
        /// Canceled
        /// </summary>
        [Description("Canceled")]
        Canceled = 12,
        /// <summary>
        /// this order status tab will be do later
        /// </summary>
        [Description("CreditRequest")]
        CreditRequest = 13
    }
}
