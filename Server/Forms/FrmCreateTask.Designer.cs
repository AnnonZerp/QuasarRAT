namespace xServer.Forms
{
    partial class FrmCreateTask
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
            this.txtName = new System.Windows.Forms.TextBox();
            this.lblName = new System.Windows.Forms.Label();
            this.lblArgs = new System.Windows.Forms.Label();
            this.txtArgs = new System.Windows.Forms.TextBox();
            this.lblPath = new System.Windows.Forms.Label();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.btnCreate = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(14, 24);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(165, 20);
            this.txtName.TabIndex = 0;
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Location = new System.Drawing.Point(11, 8);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(60, 13);
            this.lblName.TabIndex = 1;
            this.lblName.Text = "Task name";
            // 
            // lblArgs
            // 
            this.lblArgs.AutoSize = true;
            this.lblArgs.Location = new System.Drawing.Point(182, 8);
            this.lblArgs.Name = "lblArgs";
            this.lblArgs.Size = new System.Drawing.Size(83, 13);
            this.lblArgs.TabIndex = 3;
            this.lblArgs.Text = "Task arguments";
            // 
            // txtArgs
            // 
            this.txtArgs.Location = new System.Drawing.Point(185, 24);
            this.txtArgs.Name = "txtArgs";
            this.txtArgs.Size = new System.Drawing.Size(165, 20);
            this.txtArgs.TabIndex = 2;
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(11, 48);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(55, 13);
            this.lblPath.TabIndex = 5;
            this.lblPath.Text = "Task path";
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(14, 64);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(336, 20);
            this.txtPath.TabIndex = 4;
            // 
            // btnCreate
            // 
            this.btnCreate.Location = new System.Drawing.Point(275, 90);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(75, 23);
            this.btnCreate.TabIndex = 6;
            this.btnCreate.Text = "Create";
            this.btnCreate.UseVisualStyleBackColor = true;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(194, 90);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // FrmCreateTask
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(360, 120);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.lblArgs);
            this.Controls.Add(this.txtArgs);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.txtName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmCreateTask";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create task";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.Label lblArgs;
        private System.Windows.Forms.TextBox txtArgs;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Button btnCreate;
        private System.Windows.Forms.Button btnCancel;
    }
}