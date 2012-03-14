
namespace DirkSarodnick.GoogleSync.Core.Data
{
    /// <summary>
    /// Defines the DataRepository class.
    /// </summary>
    public class DataRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DataRepository"/> class.
        /// </summary>
        public DataRepository()
        {
            this.GoogleData = new GoogleData();
            this.OutlookData = new OutlookData();
        }

        /// <summary>
        /// Gets or sets the google data.
        /// </summary>
        /// <value>The google data.</value>
        public GoogleData GoogleData { get; set; }

        /// <summary>
        /// Gets or sets the outlook data.
        /// </summary>
        /// <value>The outlook data.</value>
        public OutlookData OutlookData { get; set; }
    }
}
