namespace Import_CMDB
{
    partial class frmImport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmImport));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.btnChonLog = new System.Windows.Forms.Button();
            this.txtFolderLog = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.grbAccount = new System.Windows.Forms.GroupBox();
            this.btnThoat = new System.Windows.Forms.Button();
            this.btnDangNhap = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtPass = new System.Windows.Forms.TextBox();
            this.lbPass = new System.Windows.Forms.Label();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnChonGrloader = new System.Windows.Forms.Button();
            this.txtFolderGrloader = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnTachRela = new System.Windows.Forms.Button();
            this.btnThoat2 = new System.Windows.Forms.Button();
            this.btnRelationship = new System.Windows.Forms.Button();
            this.tbnImportCI = new System.Windows.Forms.Button();
            this.txtThongBao = new System.Windows.Forms.Label();
            this.btnTachFile = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.cmbSheet = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.BrowseFile = new System.Windows.Forms.Button();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.grbAccount.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(633, 340);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.btnChonLog);
            this.tabPage1.Controls.Add(this.txtFolderLog);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.grbAccount);
            this.tabPage1.Controls.Add(this.btnChonGrloader);
            this.tabPage1.Controls.Add(this.txtFolderGrloader);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(625, 311);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Cấu hình ";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(440, 74);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(20, 13);
            this.label1.TabIndex = 43;
            this.label1.Text = "(*)";
            // 
            // btnChonLog
            // 
            this.btnChonLog.Location = new System.Drawing.Point(477, 66);
            this.btnChonLog.Name = "btnChonLog";
            this.btnChonLog.Size = new System.Drawing.Size(107, 27);
            this.btnChonLog.TabIndex = 42;
            this.btnChonLog.Text = "Chọn thư mục";
            this.btnChonLog.UseVisualStyleBackColor = true;
            this.btnChonLog.Click += new System.EventHandler(this.btnChonLog_Click);
            // 
            // txtFolderLog
            // 
            this.txtFolderLog.Enabled = false;
            this.txtFolderLog.Location = new System.Drawing.Point(187, 68);
            this.txtFolderLog.Name = "txtFolderLog";
            this.txtFolderLog.Size = new System.Drawing.Size(247, 23);
            this.txtFolderLog.TabIndex = 40;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(72, 70);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(109, 17);
            this.label3.TabIndex = 41;
            this.label3.Text = "Thư mục ghi log";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(440, 29);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(20, 13);
            this.label5.TabIndex = 39;
            this.label5.Text = "(*)";
            // 
            // grbAccount
            // 
            this.grbAccount.Controls.Add(this.btnThoat);
            this.grbAccount.Controls.Add(this.btnDangNhap);
            this.grbAccount.Controls.Add(this.label7);
            this.grbAccount.Controls.Add(this.label6);
            this.grbAccount.Controls.Add(this.txtPass);
            this.grbAccount.Controls.Add(this.lbPass);
            this.grbAccount.Controls.Add(this.txtUser);
            this.grbAccount.Controls.Add(this.label4);
            this.grbAccount.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grbAccount.Location = new System.Drawing.Point(10, 115);
            this.grbAccount.Name = "grbAccount";
            this.grbAccount.Size = new System.Drawing.Size(574, 153);
            this.grbAccount.TabIndex = 33;
            this.grbAccount.TabStop = false;
            this.grbAccount.Text = "Account Service Desk";
            // 
            // btnThoat
            // 
            this.btnThoat.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnThoat.Location = new System.Drawing.Point(284, 101);
            this.btnThoat.Name = "btnThoat";
            this.btnThoat.Size = new System.Drawing.Size(109, 30);
            this.btnThoat.TabIndex = 32;
            this.btnThoat.Text = "Thoát";
            this.btnThoat.UseVisualStyleBackColor = true;
            this.btnThoat.Click += new System.EventHandler(this.btnThoat_Click);
            // 
            // btnDangNhap
            // 
            this.btnDangNhap.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDangNhap.Location = new System.Drawing.Point(160, 101);
            this.btnDangNhap.Name = "btnDangNhap";
            this.btnDangNhap.Size = new System.Drawing.Size(109, 30);
            this.btnDangNhap.TabIndex = 31;
            this.btnDangNhap.Text = "Đăng nhập";
            this.btnDangNhap.UseVisualStyleBackColor = true;
            this.btnDangNhap.Click += new System.EventHandler(this.btnDangNhap_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.Red;
            this.label7.Location = new System.Drawing.Point(347, 67);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(20, 13);
            this.label7.TabIndex = 30;
            this.label7.Text = "(*)";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.Red;
            this.label6.Location = new System.Drawing.Point(347, 41);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(20, 13);
            this.label6.TabIndex = 29;
            this.label6.Text = "(*)";
            // 
            // txtPass
            // 
            this.txtPass.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPass.Location = new System.Drawing.Point(160, 64);
            this.txtPass.Name = "txtPass";
            this.txtPass.PasswordChar = '*';
            this.txtPass.Size = new System.Drawing.Size(181, 23);
            this.txtPass.TabIndex = 26;
            // 
            // lbPass
            // 
            this.lbPass.AutoSize = true;
            this.lbPass.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbPass.Location = new System.Drawing.Point(119, 67);
            this.lbPass.Name = "lbPass";
            this.lbPass.Size = new System.Drawing.Size(43, 17);
            this.lbPass.TabIndex = 27;
            this.lbPass.Text = "Pass:";
            // 
            // txtUser
            // 
            this.txtUser.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUser.Location = new System.Drawing.Point(160, 38);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(181, 23);
            this.txtUser.TabIndex = 24;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(120, 41);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(42, 17);
            this.label4.TabIndex = 25;
            this.label4.Text = "User:";
            // 
            // btnChonGrloader
            // 
            this.btnChonGrloader.Location = new System.Drawing.Point(477, 20);
            this.btnChonGrloader.Name = "btnChonGrloader";
            this.btnChonGrloader.Size = new System.Drawing.Size(107, 34);
            this.btnChonGrloader.TabIndex = 38;
            this.btnChonGrloader.Text = "Chọn thư mục";
            this.btnChonGrloader.UseVisualStyleBackColor = true;
            this.btnChonGrloader.Click += new System.EventHandler(this.btnChonGrloader_Click);
            // 
            // txtFolderGrloader
            // 
            this.txtFolderGrloader.Enabled = false;
            this.txtFolderGrloader.Location = new System.Drawing.Point(187, 26);
            this.txtFolderGrloader.Name = "txtFolderGrloader";
            this.txtFolderGrloader.Size = new System.Drawing.Size(247, 23);
            this.txtFolderGrloader.TabIndex = 36;
            this.txtFolderGrloader.Text = "E:\\Grloader_Setup";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(7, 29);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(174, 17);
            this.label2.TabIndex = 37;
            this.label2.Text = "Thư mục cài đặt GRloader";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btnTachRela);
            this.tabPage2.Controls.Add(this.btnThoat2);
            this.tabPage2.Controls.Add(this.btnRelationship);
            this.tabPage2.Controls.Add(this.tbnImportCI);
            this.tabPage2.Controls.Add(this.txtThongBao);
            this.tabPage2.Controls.Add(this.btnTachFile);
            this.tabPage2.Controls.Add(this.label11);
            this.tabPage2.Controls.Add(this.cmbSheet);
            this.tabPage2.Controls.Add(this.label9);
            this.tabPage2.Controls.Add(this.BrowseFile);
            this.tabPage2.Controls.Add(this.txtFile);
            this.tabPage2.Controls.Add(this.label10);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(625, 311);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Import CI";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnTachRela
            // 
            this.btnTachRela.Location = new System.Drawing.Point(222, 153);
            this.btnTachRela.Name = "btnTachRela";
            this.btnTachRela.Size = new System.Drawing.Size(137, 33);
            this.btnTachRela.TabIndex = 47;
            this.btnTachRela.Text = "Tách Relationship";
            this.btnTachRela.UseVisualStyleBackColor = true;
            this.btnTachRela.Click += new System.EventHandler(this.btnTachRela_Click);
            // 
            // btnThoat2
            // 
            this.btnThoat2.Location = new System.Drawing.Point(542, 154);
            this.btnThoat2.Name = "btnThoat2";
            this.btnThoat2.Size = new System.Drawing.Size(75, 35);
            this.btnThoat2.TabIndex = 46;
            this.btnThoat2.Text = "Thoát";
            this.btnThoat2.UseVisualStyleBackColor = true;
            this.btnThoat2.Click += new System.EventHandler(this.btnThoat_Click);
            // 
            // btnRelationship
            // 
            this.btnRelationship.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRelationship.Location = new System.Drawing.Point(370, 155);
            this.btnRelationship.Name = "btnRelationship";
            this.btnRelationship.Size = new System.Drawing.Size(148, 34);
            this.btnRelationship.TabIndex = 44;
            this.btnRelationship.Text = "Import Relationship";
            this.btnRelationship.UseVisualStyleBackColor = true;
            this.btnRelationship.Click += new System.EventHandler(this.btnRelationship_Click);
            // 
            // tbnImportCI
            // 
            this.tbnImportCI.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbnImportCI.Location = new System.Drawing.Point(129, 152);
            this.tbnImportCI.Name = "tbnImportCI";
            this.tbnImportCI.Size = new System.Drawing.Size(75, 34);
            this.tbnImportCI.TabIndex = 43;
            this.tbnImportCI.Text = "Import CI";
            this.tbnImportCI.UseVisualStyleBackColor = true;
            this.tbnImportCI.Click += new System.EventHandler(this.tbnImportCI_Click);
            // 
            // txtThongBao
            // 
            this.txtThongBao.AutoSize = true;
            this.txtThongBao.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtThongBao.ForeColor = System.Drawing.Color.Red;
            this.txtThongBao.Location = new System.Drawing.Point(42, 182);
            this.txtThongBao.Name = "txtThongBao";
            this.txtThongBao.Size = new System.Drawing.Size(0, 17);
            this.txtThongBao.TabIndex = 42;
            // 
            // btnTachFile
            // 
            this.btnTachFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTachFile.Location = new System.Drawing.Point(24, 152);
            this.btnTachFile.Name = "btnTachFile";
            this.btnTachFile.Size = new System.Drawing.Size(82, 34);
            this.btnTachFile.TabIndex = 41;
            this.btnTachFile.Text = "Tách File";
            this.btnTachFile.UseVisualStyleBackColor = true;
            this.btnTachFile.Click += new System.EventHandler(this.btnTachFile_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.Red;
            this.label11.Location = new System.Drawing.Point(429, 54);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(20, 13);
            this.label11.TabIndex = 40;
            this.label11.Text = "(*)";
            // 
            // cmbSheet
            // 
            this.cmbSheet.FormattingEnabled = true;
            this.cmbSheet.Location = new System.Drawing.Point(100, 89);
            this.cmbSheet.Name = "cmbSheet";
            this.cmbSheet.Size = new System.Drawing.Size(259, 24);
            this.cmbSheet.TabIndex = 18;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(49, 93);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(45, 17);
            this.label9.TabIndex = 17;
            this.label9.Text = "Sheet";
            // 
            // BrowseFile
            // 
            this.BrowseFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BrowseFile.Location = new System.Drawing.Point(468, 45);
            this.BrowseFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.BrowseFile.Name = "BrowseFile";
            this.BrowseFile.Size = new System.Drawing.Size(96, 28);
            this.BrowseFile.TabIndex = 16;
            this.BrowseFile.Text = "Browse";
            this.BrowseFile.UseVisualStyleBackColor = true;
            this.BrowseFile.Click += new System.EventHandler(this.BrowseFile_Click);
            // 
            // txtFile
            // 
            this.txtFile.Location = new System.Drawing.Point(100, 50);
            this.txtFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(323, 23);
            this.txtFile.TabIndex = 15;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(21, 52);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(73, 17);
            this.label10.TabIndex = 14;
            this.label10.Text = "File Import";
            // 
            // frmImport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(633, 340);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmImport";
            this.Text = "Phần mềm hỗ trợ CMDB";
            this.Load += new System.EventHandler(this.frmImport_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.grbAccount.ResumeLayout(false);
            this.grbAccount.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnChonLog;
        private System.Windows.Forms.TextBox txtFolderLog;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox grbAccount;
        private System.Windows.Forms.Button btnThoat;
        private System.Windows.Forms.Button btnDangNhap;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtPass;
        private System.Windows.Forms.Label lbPass;
        private System.Windows.Forms.TextBox txtUser;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnChonGrloader;
        private System.Windows.Forms.TextBox txtFolderGrloader;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cmbSheet;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button BrowseFile;
        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button btnTachFile;
        private System.Windows.Forms.Label txtThongBao;
        private System.Windows.Forms.Button tbnImportCI;
        private System.Windows.Forms.Button btnRelationship;
        private System.Windows.Forms.Button btnThoat2;
        private System.Windows.Forms.Button btnTachRela;
    }
}

