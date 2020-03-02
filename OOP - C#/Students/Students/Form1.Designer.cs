namespace Students
{
    partial class StudentsForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.studentsWindowMenu = new System.Windows.Forms.MenuStrip();
            this.fileMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.newStudentListMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.openStudentListMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.saveStudentListMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.viewMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.prevView = new System.Windows.Forms.ToolStripMenuItem();
            this.nextView = new System.Windows.Forms.ToolStripMenuItem();
            this.studentsMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.addStudentMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteStudentMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.nameTextBox = new System.Windows.Forms.TextBox();
            this.nameLabel = new System.Windows.Forms.Label();
            this.surnameLabel = new System.Windows.Forms.Label();
            this.surnameTextBox = new System.Windows.Forms.TextBox();
            this.facultyLabel = new System.Windows.Forms.Label();
            this.facultyTextBox = new System.Windows.Forms.TextBox();
            this.nextButton = new System.Windows.Forms.Button();
            this.prevButton = new System.Windows.Forms.Button();
            this.searchLabel = new System.Windows.Forms.Label();
            this.searchComboBox = new System.Windows.Forms.ComboBox();
            this.equalLabel = new System.Windows.Forms.Label();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.studentsWindowMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // studentsWindowMenu
            // 
            this.studentsWindowMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileMenu,
            this.viewMenu,
            this.studentsMenu});
            this.studentsWindowMenu.Location = new System.Drawing.Point(0, 0);
            this.studentsWindowMenu.Name = "studentsWindowMenu";
            this.studentsWindowMenu.Size = new System.Drawing.Size(414, 24);
            this.studentsWindowMenu.TabIndex = 0;
            this.studentsWindowMenu.Text = "menuStrip1";
            // 
            // fileMenu
            // 
            this.fileMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newStudentListMenu,
            this.openStudentListMenu,
            this.saveStudentListMenu});
            this.fileMenu.Name = "fileMenu";
            this.fileMenu.Size = new System.Drawing.Size(48, 20);
            this.fileMenu.Text = "Файл";
            // 
            // newStudentListMenu
            // 
            this.newStudentListMenu.Name = "newStudentListMenu";
            this.newStudentListMenu.Size = new System.Drawing.Size(255, 22);
            this.newStudentListMenu.Text = "Создать новый список студентов";
            this.newStudentListMenu.Click += new System.EventHandler(this.newStudentListMenu_Click);
            // 
            // openStudentListMenu
            // 
            this.openStudentListMenu.Name = "openStudentListMenu";
            this.openStudentListMenu.Size = new System.Drawing.Size(255, 22);
            this.openStudentListMenu.Text = "Открыть список студентов";
            this.openStudentListMenu.Click += new System.EventHandler(this.openStudentListMenu_Click);
            // 
            // saveStudentListMenu
            // 
            this.saveStudentListMenu.Name = "saveStudentListMenu";
            this.saveStudentListMenu.Size = new System.Drawing.Size(255, 22);
            this.saveStudentListMenu.Text = "Сохранить список студентов";
            this.saveStudentListMenu.Click += new System.EventHandler(this.saveStudentListMenu_Click);
            // 
            // viewMenu
            // 
            this.viewMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.prevView,
            this.nextView});
            this.viewMenu.Name = "viewMenu";
            this.viewMenu.Size = new System.Drawing.Size(76, 20);
            this.viewMenu.Text = "Просмотр";
            // 
            // prevView
            // 
            this.prevView.Name = "prevView";
            this.prevView.Size = new System.Drawing.Size(148, 22);
            this.prevView.Text = "Предыдущий";
            this.prevView.Click += new System.EventHandler(this.prevView_Click);
            // 
            // nextView
            // 
            this.nextView.Name = "nextView";
            this.nextView.Size = new System.Drawing.Size(148, 22);
            this.nextView.Text = "Следующий";
            this.nextView.Click += new System.EventHandler(this.nextView_Click);
            // 
            // studentsMenu
            // 
            this.studentsMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addStudentMenu,
            this.deleteStudentMenu});
            this.studentsMenu.Name = "studentsMenu";
            this.studentsMenu.Size = new System.Drawing.Size(71, 20);
            this.studentsMenu.Text = "Студенты";
            // 
            // addStudentMenu
            // 
            this.addStudentMenu.Name = "addStudentMenu";
            this.addStudentMenu.Size = new System.Drawing.Size(223, 22);
            this.addStudentMenu.Text = "Добавить нового студента";
            this.addStudentMenu.Click += new System.EventHandler(this.addStudentMenu_Click);
            // 
            // deleteStudentMenu
            // 
            this.deleteStudentMenu.Name = "deleteStudentMenu";
            this.deleteStudentMenu.Size = new System.Drawing.Size(223, 22);
            this.deleteStudentMenu.Text = "Удалить текущего студента";
            this.deleteStudentMenu.Click += new System.EventHandler(this.deleteStudentMenu_Click);
            // 
            // nameTextBox
            // 
            this.nameTextBox.Location = new System.Drawing.Point(132, 37);
            this.nameTextBox.Name = "nameTextBox";
            this.nameTextBox.Size = new System.Drawing.Size(270, 20);
            this.nameTextBox.TabIndex = 1;
            this.nameTextBox.TextChanged += new System.EventHandler(this.nameTextBox_TextChanged);
            // 
            // nameLabel
            // 
            this.nameLabel.AutoSize = true;
            this.nameLabel.Location = new System.Drawing.Point(39, 40);
            this.nameLabel.Name = "nameLabel";
            this.nameLabel.Size = new System.Drawing.Size(29, 13);
            this.nameLabel.TabIndex = 2;
            this.nameLabel.Text = "Имя";
            // 
            // surnameLabel
            // 
            this.surnameLabel.AutoSize = true;
            this.surnameLabel.Location = new System.Drawing.Point(39, 80);
            this.surnameLabel.Name = "surnameLabel";
            this.surnameLabel.Size = new System.Drawing.Size(56, 13);
            this.surnameLabel.TabIndex = 3;
            this.surnameLabel.Text = "Фамилия";
            // 
            // surnameTextBox
            // 
            this.surnameTextBox.Location = new System.Drawing.Point(132, 77);
            this.surnameTextBox.Name = "surnameTextBox";
            this.surnameTextBox.Size = new System.Drawing.Size(270, 20);
            this.surnameTextBox.TabIndex = 4;
            this.surnameTextBox.TextChanged += new System.EventHandler(this.surnameTextBox_TextChanged);
            // 
            // facultyLabel
            // 
            this.facultyLabel.AutoSize = true;
            this.facultyLabel.Location = new System.Drawing.Point(39, 120);
            this.facultyLabel.Name = "facultyLabel";
            this.facultyLabel.Size = new System.Drawing.Size(63, 13);
            this.facultyLabel.TabIndex = 5;
            this.facultyLabel.Text = "Факультет";
            // 
            // facultyTextBox
            // 
            this.facultyTextBox.Location = new System.Drawing.Point(132, 117);
            this.facultyTextBox.Name = "facultyTextBox";
            this.facultyTextBox.Size = new System.Drawing.Size(270, 20);
            this.facultyTextBox.TabIndex = 6;
            this.facultyTextBox.TextChanged += new System.EventHandler(this.facultyTextBox_TextChanged);
            // 
            // nextButton
            // 
            this.nextButton.Enabled = false;
            this.nextButton.Location = new System.Drawing.Point(312, 157);
            this.nextButton.Name = "nextButton";
            this.nextButton.Size = new System.Drawing.Size(90, 23);
            this.nextButton.TabIndex = 7;
            this.nextButton.Text = "Следующий";
            this.nextButton.UseVisualStyleBackColor = true;
            this.nextButton.Click += new System.EventHandler(this.nextButton_Click);
            // 
            // prevButton
            // 
            this.prevButton.Enabled = false;
            this.prevButton.Location = new System.Drawing.Point(12, 157);
            this.prevButton.Name = "prevButton";
            this.prevButton.Size = new System.Drawing.Size(90, 23);
            this.prevButton.TabIndex = 8;
            this.prevButton.Text = "Предыдущий";
            this.prevButton.UseVisualStyleBackColor = true;
            this.prevButton.Click += new System.EventHandler(this.prevButton_Click);
            // 
            // searchLabel
            // 
            this.searchLabel.AutoSize = true;
            this.searchLabel.Location = new System.Drawing.Point(39, 191);
            this.searchLabel.Name = "searchLabel";
            this.searchLabel.Size = new System.Drawing.Size(39, 13);
            this.searchLabel.TabIndex = 9;
            this.searchLabel.Text = "Поиск";
            // 
            // searchComboBox
            // 
            this.searchComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.searchComboBox.FormattingEnabled = true;
            this.searchComboBox.Items.AddRange(new object[] {
            "Имя",
            "Фамилия",
            "Факультет"});
            this.searchComboBox.Location = new System.Drawing.Point(12, 207);
            this.searchComboBox.Name = "searchComboBox";
            this.searchComboBox.Size = new System.Drawing.Size(90, 21);
            this.searchComboBox.TabIndex = 10;
            this.searchComboBox.SelectedIndexChanged += new System.EventHandler(this.searchComboBox_SelectedIndexChanged);
            // 
            // equalLabel
            // 
            this.equalLabel.AutoSize = true;
            this.equalLabel.Location = new System.Drawing.Point(113, 211);
            this.equalLabel.Name = "equalLabel";
            this.equalLabel.Size = new System.Drawing.Size(13, 13);
            this.equalLabel.TabIndex = 11;
            this.equalLabel.Text = "=";
            // 
            // searchTextBox
            // 
            this.searchTextBox.Location = new System.Drawing.Point(132, 208);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(270, 20);
            this.searchTextBox.TabIndex = 12;
            this.searchTextBox.TextChanged += new System.EventHandler(this.searchTextBox_TextChanged);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            this.openFileDialog.Filter = "xml-файлы|*.xml";
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.Filter = "xml-файлы|*.xml";
            // 
            // StudentsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(414, 244);
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.equalLabel);
            this.Controls.Add(this.searchComboBox);
            this.Controls.Add(this.searchLabel);
            this.Controls.Add(this.prevButton);
            this.Controls.Add(this.nextButton);
            this.Controls.Add(this.nameTextBox);
            this.Controls.Add(this.facultyTextBox);
            this.Controls.Add(this.facultyLabel);
            this.Controls.Add(this.surnameTextBox);
            this.Controls.Add(this.surnameLabel);
            this.Controls.Add(this.nameLabel);
            this.Controls.Add(this.studentsWindowMenu);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MainMenuStrip = this.studentsWindowMenu;
            this.Name = "StudentsForm";
            this.Text = "Студенты";
            this.studentsWindowMenu.ResumeLayout(false);
            this.studentsWindowMenu.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip studentsWindowMenu;
        private System.Windows.Forms.ToolStripMenuItem fileMenu;
        private System.Windows.Forms.ToolStripMenuItem viewMenu;
        private System.Windows.Forms.ToolStripMenuItem prevView;
        private System.Windows.Forms.ToolStripMenuItem nextView;
        private System.Windows.Forms.ToolStripMenuItem studentsMenu;
        private System.Windows.Forms.TextBox nameTextBox;
        private System.Windows.Forms.Label nameLabel;
        private System.Windows.Forms.Label surnameLabel;
        private System.Windows.Forms.TextBox surnameTextBox;
        private System.Windows.Forms.Label facultyLabel;
        private System.Windows.Forms.TextBox facultyTextBox;
        private System.Windows.Forms.Button nextButton;
        private System.Windows.Forms.Button prevButton;
        private System.Windows.Forms.Label searchLabel;
        private System.Windows.Forms.ComboBox searchComboBox;
        private System.Windows.Forms.Label equalLabel;
        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.ToolStripMenuItem newStudentListMenu;
        private System.Windows.Forms.ToolStripMenuItem openStudentListMenu;
        private System.Windows.Forms.ToolStripMenuItem saveStudentListMenu;
        private System.Windows.Forms.ToolStripMenuItem addStudentMenu;
        private System.Windows.Forms.ToolStripMenuItem deleteStudentMenu;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
    }
}

