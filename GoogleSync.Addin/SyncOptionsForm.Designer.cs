namespace DirkSarodnick.GoogleSync.Addin
{
    partial class SyncOptionsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.gbGoogle = new System.Windows.Forms.GroupBox();
            this.lblUsername = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.lblPassword = new System.Windows.Forms.Label();
            this.gbOptions = new System.Windows.Forms.GroupBox();
            this.rbContactMergeOutlook = new System.Windows.Forms.RadioButton();
            this.rbContactMergeGoogle = new System.Windows.Forms.RadioButton();
            this.rbContactMergeAuto = new System.Windows.Forms.RadioButton();
            this.cbIncludeContactsWithoutEmail = new System.Windows.Forms.CheckBox();
            this.gbGoogle.SuspendLayout();
            this.gbOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(297, 234);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.Location = new System.Drawing.Point(210, 234);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // gbGoogle
            // 
            this.gbGoogle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbGoogle.AutoSize = true;
            this.gbGoogle.Controls.Add(this.lblUsername);
            this.gbGoogle.Controls.Add(this.txtPassword);
            this.gbGoogle.Controls.Add(this.txtUsername);
            this.gbGoogle.Controls.Add(this.lblPassword);
            this.gbGoogle.Location = new System.Drawing.Point(12, 12);
            this.gbGoogle.Margin = new System.Windows.Forms.Padding(15);
            this.gbGoogle.Name = "gbGoogle";
            this.gbGoogle.Size = new System.Drawing.Size(360, 91);
            this.gbGoogle.TabIndex = 4;
            this.gbGoogle.TabStop = false;
            this.gbGoogle.Text = "Google Account";
            // 
            // lblUsername
            // 
            this.lblUsername.AutoSize = true;
            this.lblUsername.Location = new System.Drawing.Point(6, 29);
            this.lblUsername.Name = "lblUsername";
            this.lblUsername.Size = new System.Drawing.Size(58, 13);
            this.lblUsername.TabIndex = 5;
            this.lblUsername.Text = "Username:";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(70, 52);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(162, 20);
            this.txtPassword.TabIndex = 7;
            this.txtPassword.UseSystemPasswordChar = true;
            // 
            // txtUsername
            // 
            this.txtUsername.Location = new System.Drawing.Point(70, 26);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(162, 20);
            this.txtUsername.TabIndex = 4;
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(8, 55);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(56, 13);
            this.lblPassword.TabIndex = 6;
            this.lblPassword.Text = "Password:";
            // 
            // gbOptions
            // 
            this.gbOptions.Controls.Add(this.rbContactMergeOutlook);
            this.gbOptions.Controls.Add(this.rbContactMergeGoogle);
            this.gbOptions.Controls.Add(this.rbContactMergeAuto);
            this.gbOptions.Controls.Add(this.cbIncludeContactsWithoutEmail);
            this.gbOptions.Location = new System.Drawing.Point(12, 111);
            this.gbOptions.Name = "gbOptions";
            this.gbOptions.Size = new System.Drawing.Size(360, 117);
            this.gbOptions.TabIndex = 5;
            this.gbOptions.TabStop = false;
            this.gbOptions.Text = "Sync Options";
            // 
            // rbContactMergeOutlook
            // 
            this.rbContactMergeOutlook.AutoSize = true;
            this.rbContactMergeOutlook.Location = new System.Drawing.Point(11, 65);
            this.rbContactMergeOutlook.Name = "rbContactMergeOutlook";
            this.rbContactMergeOutlook.Size = new System.Drawing.Size(108, 17);
            this.rbContactMergeOutlook.TabIndex = 3;
            this.rbContactMergeOutlook.TabStop = true;
            this.rbContactMergeOutlook.Text = "Outlook > Google";
            this.rbContactMergeOutlook.UseVisualStyleBackColor = true;
            // 
            // rbContactMergeGoogle
            // 
            this.rbContactMergeGoogle.AutoSize = true;
            this.rbContactMergeGoogle.Location = new System.Drawing.Point(11, 42);
            this.rbContactMergeGoogle.Name = "rbContactMergeGoogle";
            this.rbContactMergeGoogle.Size = new System.Drawing.Size(108, 17);
            this.rbContactMergeGoogle.TabIndex = 2;
            this.rbContactMergeGoogle.TabStop = true;
            this.rbContactMergeGoogle.Text = "Google > Outlook";
            this.rbContactMergeGoogle.UseVisualStyleBackColor = true;
            // 
            // rbContactMergeAuto
            // 
            this.rbContactMergeAuto.AutoSize = true;
            this.rbContactMergeAuto.Location = new System.Drawing.Point(11, 19);
            this.rbContactMergeAuto.Name = "rbContactMergeAuto";
            this.rbContactMergeAuto.Size = new System.Drawing.Size(72, 17);
            this.rbContactMergeAuto.TabIndex = 1;
            this.rbContactMergeAuto.TabStop = true;
            this.rbContactMergeAuto.Text = "Automatic";
            this.rbContactMergeAuto.UseVisualStyleBackColor = true;
            // 
            // cbIncludeContactsWithoutEmail
            // 
            this.cbIncludeContactsWithoutEmail.AutoSize = true;
            this.cbIncludeContactsWithoutEmail.Location = new System.Drawing.Point(11, 88);
            this.cbIncludeContactsWithoutEmail.Name = "cbIncludeContactsWithoutEmail";
            this.cbIncludeContactsWithoutEmail.Size = new System.Drawing.Size(186, 17);
            this.cbIncludeContactsWithoutEmail.TabIndex = 0;
            this.cbIncludeContactsWithoutEmail.Text = "Include Contacts without an Email";
            this.cbIncludeContactsWithoutEmail.UseVisualStyleBackColor = true;
            // 
            // SyncOptionsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(384, 269);
            this.Controls.Add(this.gbOptions);
            this.Controls.Add(this.gbGoogle);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.MinimumSize = new System.Drawing.Size(400, 250);
            this.Name = "SyncOptionsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "GoogleSync Options";
            this.Load += new System.EventHandler(this.SyncOptions_Load);
            this.gbGoogle.ResumeLayout(false);
            this.gbGoogle.PerformLayout();
            this.gbOptions.ResumeLayout(false);
            this.gbOptions.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.GroupBox gbGoogle;
        private System.Windows.Forms.Label lblUsername;
        private System.Windows.Forms.TextBox txtUsername;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.GroupBox gbOptions;
        private System.Windows.Forms.CheckBox cbIncludeContactsWithoutEmail;
        private System.Windows.Forms.RadioButton rbContactMergeOutlook;
        private System.Windows.Forms.RadioButton rbContactMergeGoogle;
        private System.Windows.Forms.RadioButton rbContactMergeAuto;
    }
}