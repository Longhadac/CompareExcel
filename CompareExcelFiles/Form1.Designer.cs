namespace CompareExcelFiles
{
    partial class Form1
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
            this.btnFile1 = new System.Windows.Forms.Button();
            this.btnFile2 = new System.Windows.Forms.Button();
            this.txbFile1 = new System.Windows.Forms.TextBox();
            this.txbFile2 = new System.Windows.Forms.TextBox();
            this.btnCompare = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnFile1
            // 
            this.btnFile1.Location = new System.Drawing.Point(439, 10);
            this.btnFile1.Name = "btnFile1";
            this.btnFile1.Size = new System.Drawing.Size(75, 23);
            this.btnFile1.TabIndex = 0;
            this.btnFile1.Text = "Open File 1";
            this.btnFile1.UseVisualStyleBackColor = true;
            this.btnFile1.Click += new System.EventHandler(this.btnFile1_Click);
            // 
            // btnFile2
            // 
            this.btnFile2.Location = new System.Drawing.Point(439, 39);
            this.btnFile2.Name = "btnFile2";
            this.btnFile2.Size = new System.Drawing.Size(75, 23);
            this.btnFile2.TabIndex = 1;
            this.btnFile2.Text = "Open File 2";
            this.btnFile2.UseVisualStyleBackColor = true;
            this.btnFile2.Click += new System.EventHandler(this.btnFile2_Click);
            // 
            // txbFile1
            // 
            this.txbFile1.Location = new System.Drawing.Point(12, 12);
            this.txbFile1.Name = "txbFile1";
            this.txbFile1.Size = new System.Drawing.Size(421, 20);
            this.txbFile1.TabIndex = 2;
            this.txbFile1.Text = "C:\\Users\\long2\\Desktop\\T10.xlsx";
            // 
            // txbFile2
            // 
            this.txbFile2.Location = new System.Drawing.Point(12, 38);
            this.txbFile2.Name = "txbFile2";
            this.txbFile2.Size = new System.Drawing.Size(421, 20);
            this.txbFile2.TabIndex = 3;
            this.txbFile2.Text = "C:\\Users\\long2\\Desktop\\T11.xlsx";
            // 
            // btnCompare
            // 
            this.btnCompare.Location = new System.Drawing.Point(78, 82);
            this.btnCompare.Name = "btnCompare";
            this.btnCompare.Size = new System.Drawing.Size(75, 23);
            this.btnCompare.TabIndex = 4;
            this.btnCompare.Text = "Compare";
            this.btnCompare.UseVisualStyleBackColor = true;
            this.btnCompare.Click += new System.EventHandler(this.btnCompare_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(548, 390);
            this.Controls.Add(this.btnCompare);
            this.Controls.Add(this.txbFile2);
            this.Controls.Add(this.txbFile1);
            this.Controls.Add(this.btnFile2);
            this.Controls.Add(this.btnFile1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnFile1;
        private System.Windows.Forms.Button btnFile2;
        private System.Windows.Forms.TextBox txbFile1;
        private System.Windows.Forms.TextBox txbFile2;
        private System.Windows.Forms.Button btnCompare;
    }
}

