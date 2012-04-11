
namespace DirkSarodnick.GoogleSync.Core.Sync
{
    using Data;

    /// <summary>
    /// Defines the SyncBase class.
    /// </summary>
    public abstract class SyncBase : ISync
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SyncBase"/> class.
        /// </summary>
        /// <param name="repository">The repository.</param>
        protected SyncBase(DataRepository repository)
        {
            this.Repository = repository;
        }

        /// <summary>
        /// Gets or sets the repository.
        /// </summary>
        /// <value>The repository.</value>
        protected DataRepository Repository { get; set; }

        /// <summary>
        /// Syncs this instance.
        /// </summary>
        public abstract void Sync();

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        public abstract void Dispose();

        /// <summary>
        /// Items the changed.
        /// </summary>
        /// <param name="item">The item.</param>
        public abstract void ItemChanged(object item);
    }
}