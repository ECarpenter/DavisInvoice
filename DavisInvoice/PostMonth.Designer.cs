namespace DavisInvoice
{
    partial class PostMonth
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
            this.postMonthData = new System.Windows.Forms.DateTimePicker();
            this.Submit = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // postMonthData
            // 
            this.postMonthData.Location = new System.Drawing.Point(62, 66);
            this.postMonthData.Name = "postMonthData";
            this.postMonthData.Size = new System.Drawing.Size(402, 31);
            this.postMonthData.TabIndex = 0;
            this.postMonthData.ValueChanged += new System.EventHandler(this.postMonthData_ValueChanged);
            // 
            // Submit
            // 
            this.Submit.Location = new System.Drawing.Point(183, 149);
            this.Submit.Name = "Submit";
            this.Submit.Size = new System.Drawing.Size(154, 46);
            this.Submit.TabIndex = 1;
            this.Submit.Text = "Submit";
            this.Submit.UseVisualStyleBackColor = true;
            this.Submit.Click += new System.EventHandler(this.Submit_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(178, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(178, 25);
            this.label1.TabIndex = 2;
            this.label1.Text = "Enter Post Month";
            // 
            // PostMonth
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(511, 207);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Submit);
            this.Controls.Add(this.postMonthData);
            this.Name = "PostMonth";
            this.Text = "PostMonth";
            this.Load += new System.EventHandler(this.PostMonth_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker postMonthData;
        private System.Windows.Forms.Button Submit;
        private System.Windows.Forms.Label label1;
    }
}