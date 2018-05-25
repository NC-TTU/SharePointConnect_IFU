namespace SharePointConnect {
    partial class SharePointUserControl {
        /// <summary> 
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary> 
        /// Erforderliche Methode für die Designerunterstützung. 
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent() {
            this.TWI_TreeView = new System.Windows.Forms.TreeView();
            this.OpenSPButton = new System.Windows.Forms.Button();
            this.ConnectButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TWI_TreeView
            // 
            this.TWI_TreeView.Location = new System.Drawing.Point(6, 36);
            this.TWI_TreeView.Name = "TWI_TreeView";
            this.TWI_TreeView.Size = new System.Drawing.Size(191, 134);
            this.TWI_TreeView.TabIndex = 0;
            // 
            // OpenSPButton
            // 
            this.OpenSPButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OpenSPButton.Location = new System.Drawing.Point(6, 7);
            this.OpenSPButton.Name = "OpenSPButton";
            this.OpenSPButton.Size = new System.Drawing.Size(75, 23);
            this.OpenSPButton.TabIndex = 1;
            this.OpenSPButton.Text = "button1";
            this.OpenSPButton.UseVisualStyleBackColor = true;
            this.OpenSPButton.Click += new System.EventHandler(this.OpenSPButton_Click_1);
            // 
            // ConnectButton
            // 
            this.ConnectButton.Location = new System.Drawing.Point(41, 70);
            this.ConnectButton.Name = "ConnectButton";
            this.ConnectButton.Size = new System.Drawing.Size(130, 38);
            this.ConnectButton.TabIndex = 2;
            this.ConnectButton.Text = "button1";
            this.ConnectButton.UseVisualStyleBackColor = true;
            this.ConnectButton.Click += new System.EventHandler(this.ConnectButton_Click);
            // 
            // SharePointUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.ConnectButton);
            this.Controls.Add(this.OpenSPButton);
            this.Controls.Add(this.TWI_TreeView);
            this.Name = "SharePointUserControl";
            this.Size = new System.Drawing.Size(200, 200);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView TWI_TreeView;
        private System.Windows.Forms.Button OpenSPButton;
        private System.Windows.Forms.Button ConnectButton;
    }
}
