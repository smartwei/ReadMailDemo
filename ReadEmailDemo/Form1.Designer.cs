namespace ReadEmailDemo
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.listboxMail = new System.Windows.Forms.ListBox();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.txtBody = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(10, 9);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(167, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "获取OutLookMail";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // listboxMail
            // 
            this.listboxMail.FormattingEnabled = true;
            this.listboxMail.ItemHeight = 12;
            this.listboxMail.Location = new System.Drawing.Point(10, 72);
            this.listboxMail.Name = "listboxMail";
            this.listboxMail.Size = new System.Drawing.Size(167, 316);
            this.listboxMail.TabIndex = 1;
            this.listboxMail.SelectedIndexChanged += new System.EventHandler(this.listboxMail_SelectedIndexChanged);
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(190, 41);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(327, 21);
            this.txtSubject.TabIndex = 2;
            // 
            // txtBody
            // 
            this.txtBody.Location = new System.Drawing.Point(190, 90);
            this.txtBody.Multiline = true;
            this.txtBody.Name = "txtBody";
            this.txtBody.Size = new System.Drawing.Size(327, 298);
            this.txtBody.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "Outlook邮件如下";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(190, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "邮件标题";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(190, 72);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "邮件内容";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 397);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtBody);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.listboxMail);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "获取Outlook邮件";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listboxMail;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.TextBox txtBody;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}

