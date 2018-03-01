namespace EventManager
{
    partial class frmEventManager
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
            this.btnNotify = new System.Windows.Forms.Button();
            this.dlgOpenFile = new System.Windows.Forms.OpenFileDialog();
            this.txtNewFile = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.uiProgressBar1 = new Janus.Windows.EditControls.UIProgressBar();
            this.SuspendLayout();
            // 
            // btnNotify
            // 
            this.btnNotify.Location = new System.Drawing.Point(218, 119);
            this.btnNotify.Name = "btnNotify";
            this.btnNotify.Size = new System.Drawing.Size(155, 43);
            this.btnNotify.TabIndex = 0;
            this.btnNotify.Text = "Notify";
            this.btnNotify.UseVisualStyleBackColor = true;
            this.btnNotify.Click += new System.EventHandler(this.btnNotify_Click);
            // 
            // txtNewFile
            // 
            this.txtNewFile.Location = new System.Drawing.Point(51, 49);
            this.txtNewFile.Name = "txtNewFile";
            this.txtNewFile.Size = new System.Drawing.Size(255, 20);
            this.txtNewFile.TabIndex = 1;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(312, 49);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(61, 20);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // uiProgressBar1
            // 
            this.uiProgressBar1.Location = new System.Drawing.Point(2, 182);
            this.uiProgressBar1.Name = "uiProgressBar1";
            this.uiProgressBar1.Size = new System.Drawing.Size(371, 23);
            this.uiProgressBar1.TabIndex = 3;
            // 
            // frmEventManager
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(379, 208);
            this.Controls.Add(this.uiProgressBar1);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtNewFile);
            this.Controls.Add(this.btnNotify);
            this.Name = "frmEventManager";
            this.Text = "Event Manager";
            this.Load += new System.EventHandler(this.frmEventManager_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnNotify;
        private System.Windows.Forms.OpenFileDialog dlgOpenFile;
        private System.Windows.Forms.TextBox txtNewFile;
        private System.Windows.Forms.Button btnBrowse;
        private Janus.Windows.EditControls.UIProgressBar uiProgressBar1;

    }
}

