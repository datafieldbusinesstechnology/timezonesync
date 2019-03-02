namespace SSTimezone
{
    partial class SSTimeZone
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tzEventLog = new System.Diagnostics.EventLog();
            ((System.ComponentModel.ISupportInitialize)(this.tzEventLog)).BeginInit();
            // 
            // tzEventLog
            // 
            this.tzEventLog.Log = "Application";
            this.tzEventLog.Source = "Timezone Sync Service";
            // 
            // SSTimeZone
            // 
            this.ServiceName = "Timezone Sync Service";
            ((System.ComponentModel.ISupportInitialize)(this.tzEventLog)).EndInit();

        }

        #endregion

        private System.Diagnostics.EventLog tzEventLog;
    }
}
