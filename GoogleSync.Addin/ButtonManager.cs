
namespace DirkSarodnick.GoogleSync.Addin
{
    using System;
    using Core.Data;
    using Core.Sync;
    using Microsoft.Office.Core;

    /// <summary>
    /// Defines the ButtonManager class.
    /// </summary>
    public static class ButtonManager
    {
        private static CommandBarButton syncButton;
        private static CommandBarButton optionsButton;
        private static SyncManager syncManager;

        /// <summary>
        /// Creates the option button.
        /// </summary>
        /// <param name="manager">The manager.</param>
        public static void CreateButtons(SyncManager manager)
        {
            RemoveButtons();
            syncManager = manager;

            var missing = Type.Missing;
            var menubar = ApplicationData.Application.ActiveExplorer().CommandBars.ActiveMenuBar;

            var newMenuBar = (CommandBarPopup)menubar.Controls.Add(MsoControlType.msoControlPopup, missing, missing, missing, false);
            if (newMenuBar != null)
            {
                newMenuBar.Caption = "GoogleSync";
                newMenuBar.Tag = "GoogleSync";

                syncButton = (CommandBarButton)newMenuBar.Controls.Add(MsoControlType.msoControlButton, missing, missing, 1, true);
                syncButton.Style = MsoButtonStyle.msoButtonCaption;
                syncButton.Caption = "Sync now!";
                syncButton.Click += syncButton_Click;

                optionsButton = (CommandBarButton)newMenuBar.Controls.Add(MsoControlType.msoControlButton, missing, missing, missing, true);
                optionsButton.Style = MsoButtonStyle.msoButtonCaption;
                optionsButton.Caption = "Options";
                optionsButton.Click += optionsButton_Click;

                newMenuBar.Visible = true;
            }
        }

        /// <summary>
        /// Removes the option button.
        /// </summary>
        public static void RemoveButtons()
        {
            var missing = Type.Missing;
            var menubar = ApplicationData.Application.ActiveExplorer().CommandBars.ActiveMenuBar.FindControl(MsoControlType.msoControlPopup, missing, "GoogleSync", true, true);
            if (menubar != null)
            {
                menubar.Delete(true);
            }

            menubar = ApplicationData.Application.ActiveExplorer().CommandBars.ActiveMenuBar.FindControl(MsoControlType.msoControlPopup, missing, "GoogleSync", true, true);
            if (menubar != null)
            {
                menubar.Delete(true);
            }
        }

        /// <summary>
        /// The Click-Event of the options button.
        /// </summary>
        /// <param name="commandBarButton">The command bar button.</param>
        /// <param name="cancelDefault">if set to <c>true</c> [cancel default].</param>
        private static void optionsButton_Click(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            var form = new SyncOptionsForm();
            form.ShowDialog();
        }

        /// <summary>
        /// The Click-Event of the sync button.
        /// </summary>
        /// <param name="commandBarButton">The command bar button.</param>
        /// <param name="cancelDefault">if set to <c>true</c> [cancel default].</param>
        private static void syncButton_Click(CommandBarButton commandBarButton, ref bool cancelDefault)
        {
            syncManager.Start();
        }
    }
}