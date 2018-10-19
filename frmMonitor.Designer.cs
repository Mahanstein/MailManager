namespace MailManager
{
    partial class frmMonitor
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
            this.lvMonitor = new System.Windows.Forms.ListView();
            this.label1 = new System.Windows.Forms.Label();
            this.colDummy = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // lvMonitor
            // 
            this.lvMonitor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvMonitor.AutoArrange = false;
            this.lvMonitor.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colDummy});
            this.lvMonitor.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.lvMonitor.LabelWrap = false;
            this.lvMonitor.Location = new System.Drawing.Point(12, 29);
            this.lvMonitor.Name = "lvMonitor";
            this.lvMonitor.Size = new System.Drawing.Size(398, 281);
            this.lvMonitor.TabIndex = 0;
            this.lvMonitor.UseCompatibleStateImageBehavior = false;
            this.lvMonitor.View = System.Windows.Forms.View.Details;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Menssagens:";
            // 
            // colDummy
            // 
            this.colDummy.Text = "";
            this.colDummy.Width = 394;
            // 
            // frmMonitor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(422, 322);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lvMonitor);
            this.Name = "frmMonitor";
            this.Text = "Mail Manager Monitor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMonitor_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView lvMonitor;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ColumnHeader colDummy;
    }
}