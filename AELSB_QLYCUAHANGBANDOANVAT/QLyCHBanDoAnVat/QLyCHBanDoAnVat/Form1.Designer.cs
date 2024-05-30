namespace QLyCHBanDoAnVat
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            lbTenCH = new Label();
            txtLoginName = new TextBox();
            txtPassword = new TextBox();
            label1 = new Label();
            label2 = new Label();
            btDangNhap = new Button();
            label3 = new Label();
            SuspendLayout();
            // 
            // lbTenCH
            // 
            lbTenCH.Location = new Point(0, 0);
            lbTenCH.Name = "lbTenCH";
            lbTenCH.Size = new Size(100, 23);
            lbTenCH.TabIndex = 6;
            // 
            // txtLoginName
            // 
            txtLoginName.Font = new Font("Segoe UI", 15F, FontStyle.Regular, GraphicsUnit.Point);
            txtLoginName.Location = new Point(761, 304);
            txtLoginName.Name = "txtLoginName";
            txtLoginName.Size = new Size(444, 61);
            txtLoginName.TabIndex = 1;
            txtLoginName.TextChanged += txtLoginName_TextChanged;
            txtLoginName.KeyPress += txtLoginName_KeyPress;
            // 
            // txtPassword
            // 
            txtPassword.Font = new Font("Segoe UI", 15F, FontStyle.Regular, GraphicsUnit.Point);
            txtPassword.Location = new Point(761, 434);
            txtPassword.Name = "txtPassword";
            txtPassword.PasswordChar = '*';
            txtPassword.Size = new Size(444, 61);
            txtPassword.TabIndex = 2;
            txtPassword.TextChanged += txtPassword_TextChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 15F, FontStyle.Regular, GraphicsUnit.Point);
            label1.Location = new Point(275, 304);
            label1.Name = "label1";
            label1.Size = new Size(266, 54);
            label1.TabIndex = 3;
            label1.Text = "Số điện thoại:";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 15F, FontStyle.Regular, GraphicsUnit.Point);
            label2.Location = new Point(275, 434);
            label2.Name = "label2";
            label2.Size = new Size(199, 54);
            label2.TabIndex = 4;
            label2.Text = "Mật khẩu:";
            // 
            // btDangNhap
            // 
            btDangNhap.Location = new Point(657, 582);
            btDangNhap.Name = "btDangNhap";
            btDangNhap.Size = new Size(150, 62);
            btDangNhap.TabIndex = 5;
            btDangNhap.Text = "Đăng nhập";
            btDangNhap.UseVisualStyleBackColor = true;
            btDangNhap.Click += btDangNhap_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 30F, FontStyle.Regular, GraphicsUnit.Point);
            label3.Location = new Point(617, 38);
            label3.Name = "label3";
            label3.Size = new Size(264, 106);
            label3.TabIndex = 7;
            label3.Text = "AELSB";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(13F, 32F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1474, 729);
            Controls.Add(label3);
            Controls.Add(btDangNhap);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(txtPassword);
            Controls.Add(txtLoginName);
            Controls.Add(lbTenCH);
            Name = "Form1";
            Text = "AELSB";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label lbTenCH;
        private TextBox txtLoginName;
        private TextBox txtPassword;
        private Label label1;
        private Label label2;
        private Button btDangNhap;
        private Label label3;
    }
}