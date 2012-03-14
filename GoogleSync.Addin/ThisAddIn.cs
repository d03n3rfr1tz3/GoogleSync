
namespace DirkSarodnick.GoogleSync.Addin
{
    using System.Reflection;
    using System.Timers;
    using Core.Data;
    using Core.Sync;

    /// <summary>
    /// Defines the ThisAddIn class.
    /// </summary>
    public partial class ThisAddIn
    {
        private bool started;
        private bool stopped;
        private readonly Timer timer = new Timer();
        private readonly SyncManager syncManager = new SyncManager();

        /// <summary>
        /// Handles the Startup event of the ThisAddIn control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (started) return;
            started = true;

            ApplicationData.Application = this.Application;
            ApplicationData.GoogleApplication = Assembly.GetExecutingAssembly().GetName().Name;
            ButtonManager.CreateButtons(syncManager);

            timer.AutoReset = true;
            timer.Enabled = true;
            timer.Interval = 1000;
            timer.Elapsed += timer_Elapsed;
            timer.Start();
        }

        /// <summary>
        /// Handles the Shutdown event of the ThisAddIn control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (stopped) return;
            stopped = true;

            syncManager.Dispose();
        }

        void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            timer.Interval = 3600000;
            syncManager.Start();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}