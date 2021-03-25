
namespace ExcelToKompas
{
    partial class Form1
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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn_select_table = new System.Windows.Forms.Button();
            this.table_path = new System.Windows.Forms.TextBox();
            this.dir_in = new System.Windows.Forms.TextBox();
            this.btn_select_dir_in = new System.Windows.Forms.Button();
            this.dir_out = new System.Windows.Forms.TextBox();
            this.btn_select_style = new System.Windows.Forms.Button();
            this.btn_select_dir_out = new System.Windows.Forms.Button();
            this.style_path = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.multiplier = new System.Windows.Forms.NumericUpDown();
            this.cb_connect_all = new System.Windows.Forms.CheckBox();
            this.btn_ok = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.subsection_name = new System.Windows.Forms.ComboBox();
            this.manufacturer_name = new System.Windows.Forms.ComboBox();
            this.code_name = new System.Windows.Forms.ComboBox();
            this.user_name = new System.Windows.Forms.ComboBox();
            this.material_name = new System.Windows.Forms.ComboBox();
            this.mass_name = new System.Windows.Forms.ComboBox();
            this.add_name = new System.Windows.Forms.ComboBox();
            this.count_name = new System.Windows.Forms.ComboBox();
            this.name_name = new System.Windows.Forms.ComboBox();
            this.mark_name = new System.Windows.Forms.ComboBox();
            this.position_name = new System.Windows.Forms.ComboBox();
            this.zone_name = new System.Windows.Forms.ComboBox();
            this.format_name = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.info_label = new System.Windows.Forms.Label();
            this.btn_upd = new System.Windows.Forms.Button();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.multiplier)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn_select_table);
            this.groupBox2.Controls.Add(this.table_path);
            this.groupBox2.Controls.Add(this.dir_in);
            this.groupBox2.Controls.Add(this.btn_select_dir_in);
            this.groupBox2.Controls.Add(this.dir_out);
            this.groupBox2.Controls.Add(this.btn_select_style);
            this.groupBox2.Controls.Add(this.btn_select_dir_out);
            this.groupBox2.Controls.Add(this.style_path);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBox2.Location = new System.Drawing.Point(12, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(626, 138);
            this.groupBox2.TabIndex = 16;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Директории и файлы";
            // 
            // btn_select_table
            // 
            this.btn_select_table.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btn_select_table.Location = new System.Drawing.Point(452, 106);
            this.btn_select_table.Name = "btn_select_table";
            this.btn_select_table.Size = new System.Drawing.Size(169, 23);
            this.btn_select_table.TabIndex = 8;
            this.btn_select_table.Text = "Выбрать таблицу разделов";
            this.btn_select_table.UseVisualStyleBackColor = true;
            this.btn_select_table.Click += new System.EventHandler(this.btn_select_table_Click);
            // 
            // table_path
            // 
            this.table_path.Enabled = false;
            this.table_path.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.table_path.Location = new System.Drawing.Point(6, 108);
            this.table_path.Name = "table_path";
            this.table_path.Size = new System.Drawing.Size(436, 19);
            this.table_path.TabIndex = 7;
            // 
            // dir_in
            // 
            this.dir_in.Enabled = false;
            this.dir_in.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dir_in.Location = new System.Drawing.Point(6, 25);
            this.dir_in.Name = "dir_in";
            this.dir_in.Size = new System.Drawing.Size(436, 19);
            this.dir_in.TabIndex = 1;
            // 
            // btn_select_dir_in
            // 
            this.btn_select_dir_in.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btn_select_dir_in.Location = new System.Drawing.Point(452, 23);
            this.btn_select_dir_in.Name = "btn_select_dir_in";
            this.btn_select_dir_in.Size = new System.Drawing.Size(169, 23);
            this.btn_select_dir_in.TabIndex = 0;
            this.btn_select_dir_in.Text = "Путь к папке с исходниками";
            this.btn_select_dir_in.UseVisualStyleBackColor = true;
            this.btn_select_dir_in.Click += new System.EventHandler(this.btn_select_dir_in_Click);
            // 
            // dir_out
            // 
            this.dir_out.Enabled = false;
            this.dir_out.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dir_out.Location = new System.Drawing.Point(6, 53);
            this.dir_out.Name = "dir_out";
            this.dir_out.Size = new System.Drawing.Size(436, 19);
            this.dir_out.TabIndex = 2;
            // 
            // btn_select_style
            // 
            this.btn_select_style.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btn_select_style.Location = new System.Drawing.Point(452, 79);
            this.btn_select_style.Name = "btn_select_style";
            this.btn_select_style.Size = new System.Drawing.Size(169, 23);
            this.btn_select_style.TabIndex = 6;
            this.btn_select_style.Text = "Выбрать файл стиля";
            this.btn_select_style.UseVisualStyleBackColor = true;
            this.btn_select_style.Click += new System.EventHandler(this.btn_select_style_Click);
            // 
            // btn_select_dir_out
            // 
            this.btn_select_dir_out.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btn_select_dir_out.Location = new System.Drawing.Point(452, 51);
            this.btn_select_dir_out.Name = "btn_select_dir_out";
            this.btn_select_dir_out.Size = new System.Drawing.Size(169, 23);
            this.btn_select_dir_out.TabIndex = 3;
            this.btn_select_dir_out.Text = "Путь к папке с результатом";
            this.btn_select_dir_out.UseVisualStyleBackColor = true;
            this.btn_select_dir_out.Click += new System.EventHandler(this.btn_select_dir_out_Click);
            // 
            // style_path
            // 
            this.style_path.Enabled = false;
            this.style_path.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.style_path.Location = new System.Drawing.Point(6, 81);
            this.style_path.Name = "style_path";
            this.style_path.Size = new System.Drawing.Size(436, 19);
            this.style_path.TabIndex = 5;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_upd);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.multiplier);
            this.groupBox1.Controls.Add(this.cb_connect_all);
            this.groupBox1.Controls.Add(this.btn_ok);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBox1.Location = new System.Drawing.Point(644, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(171, 138);
            this.groupBox1.TabIndex = 17;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Параметры";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(27, 53);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 17;
            this.label1.Text = "Множитель:";
            // 
            // multiplier
            // 
            this.multiplier.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.multiplier.Location = new System.Drawing.Point(101, 51);
            this.multiplier.Margin = new System.Windows.Forms.Padding(2);
            this.multiplier.Name = "multiplier";
            this.multiplier.Size = new System.Drawing.Size(37, 19);
            this.multiplier.TabIndex = 16;
            this.multiplier.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // cb_connect_all
            // 
            this.cb_connect_all.AutoSize = true;
            this.cb_connect_all.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.cb_connect_all.Location = new System.Drawing.Point(12, 27);
            this.cb_connect_all.Name = "cb_connect_all";
            this.cb_connect_all.Size = new System.Drawing.Size(145, 17);
            this.cb_connect_all.TabIndex = 14;
            this.cb_connect_all.Text = "Соединить в один файл";
            this.cb_connect_all.UseVisualStyleBackColor = true;
            // 
            // btn_ok
            // 
            this.btn_ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btn_ok.Location = new System.Drawing.Point(6, 106);
            this.btn_ok.Name = "btn_ok";
            this.btn_ok.Size = new System.Drawing.Size(159, 23);
            this.btn_ok.TabIndex = 4;
            this.btn_ok.Text = "Конвертировать";
            this.btn_ok.UseVisualStyleBackColor = true;
            this.btn_ok.Click += new System.EventHandler(this.btn_ok_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.subsection_name);
            this.groupBox3.Controls.Add(this.manufacturer_name);
            this.groupBox3.Controls.Add(this.code_name);
            this.groupBox3.Controls.Add(this.user_name);
            this.groupBox3.Controls.Add(this.material_name);
            this.groupBox3.Controls.Add(this.mass_name);
            this.groupBox3.Controls.Add(this.add_name);
            this.groupBox3.Controls.Add(this.count_name);
            this.groupBox3.Controls.Add(this.name_name);
            this.groupBox3.Controls.Add(this.mark_name);
            this.groupBox3.Controls.Add(this.position_name);
            this.groupBox3.Controls.Add(this.zone_name);
            this.groupBox3.Controls.Add(this.format_name);
            this.groupBox3.Controls.Add(this.label14);
            this.groupBox3.Controls.Add(this.label13);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Location = new System.Drawing.Point(821, 11);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(204, 362);
            this.groupBox3.TabIndex = 23;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Имена колонок";
            // 
            // subsection_name
            // 
            this.subsection_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.subsection_name.FormattingEnabled = true;
            this.subsection_name.Location = new System.Drawing.Point(95, 327);
            this.subsection_name.Margin = new System.Windows.Forms.Padding(2);
            this.subsection_name.Name = "subsection_name";
            this.subsection_name.Size = new System.Drawing.Size(104, 21);
            this.subsection_name.TabIndex = 59;
            // 
            // manufacturer_name
            // 
            this.manufacturer_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.manufacturer_name.FormattingEnabled = true;
            this.manufacturer_name.Location = new System.Drawing.Point(95, 302);
            this.manufacturer_name.Margin = new System.Windows.Forms.Padding(2);
            this.manufacturer_name.Name = "manufacturer_name";
            this.manufacturer_name.Size = new System.Drawing.Size(104, 21);
            this.manufacturer_name.TabIndex = 58;
            // 
            // code_name
            // 
            this.code_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.code_name.FormattingEnabled = true;
            this.code_name.Location = new System.Drawing.Point(95, 277);
            this.code_name.Margin = new System.Windows.Forms.Padding(2);
            this.code_name.Name = "code_name";
            this.code_name.Size = new System.Drawing.Size(104, 21);
            this.code_name.TabIndex = 57;
            // 
            // user_name
            // 
            this.user_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.user_name.FormattingEnabled = true;
            this.user_name.Location = new System.Drawing.Point(95, 252);
            this.user_name.Margin = new System.Windows.Forms.Padding(2);
            this.user_name.Name = "user_name";
            this.user_name.Size = new System.Drawing.Size(104, 21);
            this.user_name.TabIndex = 56;
            // 
            // material_name
            // 
            this.material_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.material_name.FormattingEnabled = true;
            this.material_name.Location = new System.Drawing.Point(95, 227);
            this.material_name.Margin = new System.Windows.Forms.Padding(2);
            this.material_name.Name = "material_name";
            this.material_name.Size = new System.Drawing.Size(104, 21);
            this.material_name.TabIndex = 55;
            // 
            // mass_name
            // 
            this.mass_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.mass_name.FormattingEnabled = true;
            this.mass_name.Location = new System.Drawing.Point(95, 202);
            this.mass_name.Margin = new System.Windows.Forms.Padding(2);
            this.mass_name.Name = "mass_name";
            this.mass_name.Size = new System.Drawing.Size(104, 21);
            this.mass_name.TabIndex = 54;
            // 
            // add_name
            // 
            this.add_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.add_name.FormattingEnabled = true;
            this.add_name.Location = new System.Drawing.Point(95, 177);
            this.add_name.Margin = new System.Windows.Forms.Padding(2);
            this.add_name.Name = "add_name";
            this.add_name.Size = new System.Drawing.Size(104, 21);
            this.add_name.TabIndex = 53;
            // 
            // count_name
            // 
            this.count_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.count_name.FormattingEnabled = true;
            this.count_name.Location = new System.Drawing.Point(95, 152);
            this.count_name.Margin = new System.Windows.Forms.Padding(2);
            this.count_name.Name = "count_name";
            this.count_name.Size = new System.Drawing.Size(104, 21);
            this.count_name.TabIndex = 52;
            // 
            // name_name
            // 
            this.name_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.name_name.FormattingEnabled = true;
            this.name_name.Location = new System.Drawing.Point(95, 127);
            this.name_name.Margin = new System.Windows.Forms.Padding(2);
            this.name_name.Name = "name_name";
            this.name_name.Size = new System.Drawing.Size(104, 21);
            this.name_name.TabIndex = 51;
            // 
            // mark_name
            // 
            this.mark_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.mark_name.FormattingEnabled = true;
            this.mark_name.Location = new System.Drawing.Point(95, 102);
            this.mark_name.Margin = new System.Windows.Forms.Padding(2);
            this.mark_name.Name = "mark_name";
            this.mark_name.Size = new System.Drawing.Size(104, 21);
            this.mark_name.TabIndex = 50;
            // 
            // position_name
            // 
            this.position_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.position_name.FormattingEnabled = true;
            this.position_name.Location = new System.Drawing.Point(95, 77);
            this.position_name.Margin = new System.Windows.Forms.Padding(2);
            this.position_name.Name = "position_name";
            this.position_name.Size = new System.Drawing.Size(104, 21);
            this.position_name.TabIndex = 49;
            // 
            // zone_name
            // 
            this.zone_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.zone_name.FormattingEnabled = true;
            this.zone_name.Location = new System.Drawing.Point(95, 52);
            this.zone_name.Margin = new System.Windows.Forms.Padding(2);
            this.zone_name.Name = "zone_name";
            this.zone_name.Size = new System.Drawing.Size(104, 21);
            this.zone_name.TabIndex = 48;
            // 
            // format_name
            // 
            this.format_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.format_name.FormattingEnabled = true;
            this.format_name.Location = new System.Drawing.Point(95, 27);
            this.format_name.Margin = new System.Windows.Forms.Padding(2);
            this.format_name.Name = "format_name";
            this.format_name.Size = new System.Drawing.Size(104, 21);
            this.format_name.TabIndex = 18;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(5, 330);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(71, 13);
            this.label14.TabIndex = 47;
            this.label14.Text = "Тип изделия";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(5, 305);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(78, 13);
            this.label13.TabIndex = 45;
            this.label13.Text = "Изготовитель";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(5, 105);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(74, 13);
            this.label5.TabIndex = 30;
            this.label5.Text = "Обозначение";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(5, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 13);
            this.label3.TabIndex = 23;
            this.label3.Text = "Позиция";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(6, 280);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(26, 13);
            this.label12.TabIndex = 43;
            this.label12.Text = "Код";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(5, 155);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(66, 13);
            this.label6.TabIndex = 29;
            this.label6.Text = "Количество";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(5, 255);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(80, 13);
            this.label11.TabIndex = 42;
            this.label11.Text = "Пользователь";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(5, 130);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(83, 13);
            this.label7.TabIndex = 28;
            this.label7.Text = "Наименование";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(5, 230);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(57, 13);
            this.label9.TabIndex = 35;
            this.label9.Text = "Материал";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(5, 180);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(70, 13);
            this.label8.TabIndex = 36;
            this.label8.Text = "Примечание";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 205);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(40, 13);
            this.label10.TabIndex = 34;
            this.label10.Text = "Масса";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(5, 30);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(49, 13);
            this.label4.TabIndex = 24;
            this.label4.Text = "Формат";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(32, 13);
            this.label2.TabIndex = 22;
            this.label2.Text = "Зона";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.listBox1);
            this.groupBox4.Location = new System.Drawing.Point(9, 156);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(806, 188);
            this.groupBox4.TabIndex = 25;
            this.groupBox4.TabStop = false;
            // 
            // listBox1
            // 
            this.listBox1.AllowDrop = true;
            this.listBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(6, 16);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(794, 160);
            this.listBox1.TabIndex = 0;
            this.listBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this.listBox1_DragDrop);
            this.listBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this.listBox1_DragEnter);
            // 
            // info_label
            // 
            this.info_label.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.info_label.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.info_label.Location = new System.Drawing.Point(15, 347);
            this.info_label.Name = "info_label";
            this.info_label.Size = new System.Drawing.Size(794, 23);
            this.info_label.TabIndex = 24;
            this.info_label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btn_upd
            // 
            this.btn_upd.Location = new System.Drawing.Point(6, 79);
            this.btn_upd.Name = "btn_upd";
            this.btn_upd.Size = new System.Drawing.Size(159, 23);
            this.btn_upd.TabIndex = 1;
            this.btn_upd.Text = "Обновить папку";
            this.btn_upd.UseVisualStyleBackColor = true;
            this.btn_upd.Click += new System.EventHandler(this.btn_upd_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1035, 383);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.info_label);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.MaximumSize = new System.Drawing.Size(1051, 422);
            this.MinimumSize = new System.Drawing.Size(1051, 422);
            this.Name = "Form1";
            this.Text = "Конвертер из Excel в Kompas 3D";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.multiplier)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btn_select_table;
        private System.Windows.Forms.TextBox table_path;
        private System.Windows.Forms.TextBox dir_in;
        private System.Windows.Forms.Button btn_select_dir_in;
        private System.Windows.Forms.TextBox dir_out;
        private System.Windows.Forms.Button btn_select_style;
        private System.Windows.Forms.Button btn_select_dir_out;
        private System.Windows.Forms.TextBox style_path;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown multiplier;
        private System.Windows.Forms.CheckBox cb_connect_all;
        private System.Windows.Forms.Button btn_ok;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ComboBox subsection_name;
        private System.Windows.Forms.ComboBox manufacturer_name;
        private System.Windows.Forms.ComboBox code_name;
        private System.Windows.Forms.ComboBox user_name;
        private System.Windows.Forms.ComboBox material_name;
        private System.Windows.Forms.ComboBox mass_name;
        private System.Windows.Forms.ComboBox add_name;
        private System.Windows.Forms.ComboBox count_name;
        private System.Windows.Forms.ComboBox name_name;
        private System.Windows.Forms.ComboBox mark_name;
        private System.Windows.Forms.ComboBox position_name;
        private System.Windows.Forms.ComboBox zone_name;
        private System.Windows.Forms.ComboBox format_name;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Label info_label;
        private System.Windows.Forms.Button btn_upd;
    }
}

