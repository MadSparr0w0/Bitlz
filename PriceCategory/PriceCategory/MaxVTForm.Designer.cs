namespace PriceCategory
{
    partial class MaxVTForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MaxVTForm));
            this.vtCB = new System.Windows.Forms.ComboBox();
            this.maxPowerTB = new System.Windows.Forms.TextBox();
            this.DoneB = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // vtCB
            // 
            this.vtCB.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(86)))), ((int)(((byte)(86)))), ((int)(((byte)(86)))));
            this.vtCB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.vtCB.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F);
            this.vtCB.FormattingEnabled = true;
            this.vtCB.Items.AddRange(new object[] {
            "кВт",
            "МВт"});
            this.vtCB.Location = new System.Drawing.Point(191, 12);
            this.vtCB.Name = "vtCB";
            this.vtCB.Size = new System.Drawing.Size(72, 37);
            this.vtCB.TabIndex = 4;
            this.vtCB.Text = "кВт";
            // 
            // maxPowerTB
            // 
            this.maxPowerTB.BackColor = System.Drawing.Color.DarkGray;
            this.maxPowerTB.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.maxPowerTB.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F);
            this.maxPowerTB.Location = new System.Drawing.Point(12, 12);
            this.maxPowerTB.Name = "maxPowerTB";
            this.maxPowerTB.Size = new System.Drawing.Size(163, 37);
            this.maxPowerTB.TabIndex = 3;
            // 
            // DoneB
            // 
            this.DoneB.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(86)))), ((int)(((byte)(86)))), ((int)(((byte)(86)))));
            this.DoneB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DoneB.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.DoneB.Location = new System.Drawing.Point(12, 55);
            this.DoneB.Name = "DoneB";
            this.DoneB.Size = new System.Drawing.Size(163, 32);
            this.DoneB.TabIndex = 5;
            this.DoneB.Text = "Готово";
            this.DoneB.UseVisualStyleBackColor = false;
            this.DoneB.Click += new System.EventHandler(this.DoneB_Click);
            // 
            // MaxVTForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(44)))), ((int)(((byte)(44)))));
            this.ClientSize = new System.Drawing.Size(275, 99);
            this.Controls.Add(this.DoneB);
            this.Controls.Add(this.vtCB);
            this.Controls.Add(this.maxPowerTB);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MaxVTForm";
            this.Text = "Максимальное потребление";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox vtCB;
        private System.Windows.Forms.TextBox maxPowerTB;
        private System.Windows.Forms.Button DoneB;
    }
}