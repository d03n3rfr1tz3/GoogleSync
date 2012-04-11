
namespace DirkSarodnick.GoogleSync.Addin
{
    using System;
    using System.Windows.Forms;
    using Core.Data;

    /// <summary>
    /// Defines the SyncOptionsForm class.
    /// </summary>
    public partial class SyncOptionsForm : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SyncOptionsForm"/> class.
        /// </summary>
        public SyncOptionsForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Handles the Load event of the SyncOptions control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void SyncOptions_Load(object sender, EventArgs e)
        {
            ApplicationData.ResolveGoogleAccount();
            this.LoadOptions();
        }

        /// <summary>
        /// Handles the Click event of the btnCancel control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Handles the Click event of the btnOK control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void btnOK_Click(object sender, EventArgs e)
        {
            this.SaveOptions();
            this.Close();
        }

        /// <summary>
        /// Loads the options.
        /// </summary>
        public void LoadOptions()
        {
            this.txtUsername.Text = ApplicationData.GoogleUsername;
            this.txtPassword.Text = ApplicationData.GooglePassword;
            this.cbIncludeContactsWithoutEmail.Checked = ApplicationData.IncludeContactWithoutEmail;
            this.rbContactMergeAuto.Checked = ApplicationData.ContactBehavior == ContactBehavior.Automatic;
            this.rbContactMergeGoogle.Checked = ApplicationData.ContactBehavior == ContactBehavior.GoogleOverOutlook;
            this.rbContactMergeOutlook.Checked = ApplicationData.ContactBehavior == ContactBehavior.OutlookOverGoogle;
        }

        /// <summary>
        /// Saves the options.
        /// </summary>
        public void SaveOptions()
        {
            ApplicationData.GoogleUsername = this.txtUsername.Text;
            ApplicationData.GooglePassword = this.txtPassword.Text;
            ApplicationData.IncludeContactWithoutEmail = this.cbIncludeContactsWithoutEmail.Checked;
            if (rbContactMergeAuto.Checked) ApplicationData.ContactBehavior = ContactBehavior.Automatic;
            if (rbContactMergeGoogle.Checked) ApplicationData.ContactBehavior = ContactBehavior.GoogleOverOutlook;
            if (rbContactMergeOutlook.Checked) ApplicationData.ContactBehavior = ContactBehavior.OutlookOverGoogle;
        }
    }
}
