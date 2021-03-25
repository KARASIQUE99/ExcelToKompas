using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Kompas6API5;
using Microsoft.WindowsAPICodePack.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToKompas
{
    public partial class Form1 : Form
    {
        private KompasObject kompas;
        private Excel.Application excelApp;
        private delegate void SafeCallDelegate(string text);

        private string inputDirName = "";
        private string outputDirName = "";
        private string styleFileName = "";

        List<List<string>> table = null;

        public Form1()
        {
            InitializeComponent();
            InitializeApplications(true);
        }

        public void reload()
        {
            listBox1.Items.Clear();

            foreach (string fileName in Directory.GetFiles(inputDirName))
                if ((fileName.EndsWith(".xlsx") || fileName.EndsWith(".xlsm") || fileName.EndsWith(".xls")) && !fileName.Contains("~"))
                    listBox1.Items.Add(fileName + ", " + 1);

            if (listBox1.Items.Count > 0)
            {

                List<string> columns = Utils.getFirstRowAsList(
                    Utils.getWorksheetByFileName(excelApp,
                    Utils.getFileNameWithoutCount(
                        Convert.ToString(listBox1.Items[0]))));

                format_name.Items.Clear();
                zone_name.Items.Clear();
                position_name.Items.Clear();
                mark_name.Items.Clear();
                name_name.Items.Clear();
                count_name.Items.Clear();
                add_name.Items.Clear();
                mass_name.Items.Clear();
                material_name.Items.Clear();
                user_name.Items.Clear();
                code_name.Items.Clear();
                manufacturer_name.Items.Clear();
                subsection_name.Items.Clear();

                format_name.Items.AddRange(columns.ToArray());
                zone_name.Items.AddRange(columns.ToArray());
                position_name.Items.AddRange(columns.ToArray());
                mark_name.Items.AddRange(columns.ToArray());
                name_name.Items.AddRange(columns.ToArray());
                count_name.Items.AddRange(columns.ToArray());
                add_name.Items.AddRange(columns.ToArray());
                mass_name.Items.AddRange(columns.ToArray());
                material_name.Items.AddRange(columns.ToArray());
                user_name.Items.AddRange(columns.ToArray());
                code_name.Items.AddRange(columns.ToArray());
                manufacturer_name.Items.AddRange(columns.ToArray());
                subsection_name.Items.AddRange(columns.ToArray());

                if (columns.ToArray().Contains("1CType"))
                    subsection_name.Text = "1CType";
                if (columns.ToArray().Contains("PartNumberFrom1C"))
                    name_name.Text = "PartNumberFrom1C";
                if (columns.ToArray().Contains("Quantity"))
                    count_name.Text = "Quantity";
                if (columns.ToArray().Contains("Designator"))
                {
                    add_name.Text = "Designator";
                    user_name.Text = "Designator";
                }

            }
        }

        public void loadPaths()
        {
            try
            {
                if (File.Exists("paths.txt"))
                {
                    string[] paths = File.ReadAllLines("paths.txt");
                    inputDirName = paths[0];
                    outputDirName = paths[1];
                    styleFileName = paths[2];

                    dir_in.Text = paths[0];
                    dir_out.Text = paths[1];
                    style_path.Text = paths[2];
                    table_path.Text = paths[3];

                    reload();

                    table = new List<List<string>>();
                    string[] lines = File.ReadAllLines(paths[3]);

                    foreach (string line in lines)
                    {
                        List<string> item = new List<string>();
                        item.Add(line.Split('=')[0]);
                        item.Add(line.Split('=')[1]);
                        item.Add(line.Split('=')[2]);
                        table.Add(item);
                    }

                }
            } catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка загрузки путей из файла paths.txt!\n\n" +
                                "Проверьте правильность путей в файле paths.txt или " +
                                "проигнорируйте ошибку и укажите пути вручную.\n\n" +
                                "Текст ошибки: " + e.Message);
                info_label.Text = "Ошибка загрузки путей из файла paths.txt!";
            }
}

        public void InitializeApplications(bool load)
        {
            try
            {
                #if __LIGHT_VERSION__
				    Type t = Type.GetTypeFromProgID("KOMPASLT.Application.5");
                #else
                    Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
                #endif

                excelApp = new Excel.Application();
                kompas = (KompasObject)Activator.CreateInstance(t);
                if(load) loadPaths();
            }
            catch (Exception ex)
            { 
                MessageBox.Show("Произошла ошибка инициализации!\n\n" +
                                "Проверьте работу ПО Компас и ПО Excel, " +
                                "наличие соответствующих лицензий.\n\n" +
                                "Текст ошибки: " + ex.Message);
                info_label.Text = "Ошибка инициализации!";
            }
            
        }

        private void btn_select_dir_in_Click(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();

                CommonOpenFileDialog FBD = new CommonOpenFileDialog();
                FBD.IsFolderPicker = true;
                FBD.Title = "Выберите папку с исходниками";

                if (FBD.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    inputDirName = FBD.FileName;
                    dir_in.Text = inputDirName;

                    if (outputDirName == "" && dir_out.Text == "")
                    {
                        outputDirName = inputDirName + @"\output";
                        dir_out.Text = outputDirName;
                    }
                }

                if (inputDirName == "") return;

                reload();

                info_label.Text = "";
            }
            catch (Exception ex)
            {
                info_label.Text = "Ошибка загрузки исходников!";
            }
        }

        private void btn_select_dir_out_Click(object sender, EventArgs e)
        {
            try
            {
                CommonOpenFileDialog FBD = new CommonOpenFileDialog();
                FBD.IsFolderPicker = true;
                FBD.Title = "Выберите папку для вывода";

                if (FBD.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    outputDirName = FBD.FileName + @"\output";
                    dir_out.Text = outputDirName;

                }
                info_label.Text = "";
            }
            catch (Exception ex)
            {
                info_label.Text = "Ошибка при выборе папки для результатов!";
            }
        }

        private void btn_select_style_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog OPD = new OpenFileDialog();
                OPD.Title = "Выберите файл стиля";
                OPD.Filter = "Файлы стиля (*.lyt)|*.lyt;";

                if (OPD.ShowDialog() == DialogResult.OK)
                {
                    styleFileName = OPD.FileName;
                    style_path.Text = styleFileName;
                }
                info_label.Text = "";
            }
            catch (Exception ex)
            {
                info_label.Text = "Ошибка загрузки файла стиля!";
            }
        }

        private void btn_select_table_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog OPD = new OpenFileDialog();
                OPD.Title = "Выберите файл с таблицей разделов";
                OPD.Filter = "txt-файлы (*.txt)|*.txt;";

                if (OPD.ShowDialog() == DialogResult.OK)
                {
                    table_path.Text = OPD.FileName;
                    table = new List<List<string>>();
                    string[] lines = File.ReadAllLines(OPD.FileName);
                    foreach (string line in lines)
                    {
                        List<string> item = new List<string>();
                        item.Add(line.Split('=')[0]);
                        item.Add(line.Split('=')[1]);
                        item.Add(line.Split('=')[2]);
                        table.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                info_label.Text = "Ошибка загрузки файла с таблицей!";
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                kompas.Quit();
                excelApp.Quit();
            }
            catch (Exception ex) { }
        }

        private Dictionary<string, List<Row>> getInputDataFromExcel(Excel.Application excelApp, string[] fileEntries, bool connect, Dictionary<string, string> columnNames)
        {
            Dictionary<string, List<Row>> outputs = new Dictionary<string, List<Row>>();

            if (connect)
            {
                List<Row> singleInput = new List<Row>();
                foreach (string fileName in fileEntries)
                {
                    int count = Convert.ToInt32(fileName.Split(' ').Last());
                    if (count < 1) continue;
                    for (int i = 0; i < count; i++)
                    {
                        string fName = Utils.getFileNameWithoutCount(fileName);
                        WriteTextSafe(fName);
                        singleInput.AddRange(Utils.GetDataFromExcel(Utils.getWorksheetByFileName(excelApp, Utils.getFileNameWithoutCount(fileName)), columnNames));
                    }

                }
                outputs.Add("result", singleInput);
            }
            else
            {
                foreach (string fileName in fileEntries)
                {
                    int count = Convert.ToInt32(fileName.Split(' ').Last());
                    if (count < 1) continue;
                    string fileNameWithoutCount = Utils.getFileNameWithoutCount(fileName);
                    WriteTextSafe(fileNameWithoutCount);
                    List<Row> rws = Utils.GetDataFromExcel(Utils.getWorksheetByFileName(excelApp, fileNameWithoutCount), columnNames);
                    List<Row> rows = new List<Row>();
                    for (int i = 0; i < count; i++)
                    {
                        rows.AddRange(rws);
                    }

                    outputs[fileNameWithoutCount] = rows;
                }
            }

            WriteTextSafe("Преобразование данных в формат КОМПАС 3D");
            return outputs;
        }

        private void runExcelToKompas(bool connect, Dictionary<string, string> d)
        {
            Dictionary<string, List<Row>> inputs = getInputDataFromExcel(excelApp, Utils.getEntities(listBox1.Items), connect, d);
            foreach (KeyValuePair<string, List<Row>> entry in inputs)
            {
                string fileName = entry.Key;
                if (entry.Key.EndsWith(".xlsx") || entry.Key.EndsWith(".xlsm")) fileName = fileName.Substring(0, fileName.Length - 5);
                Utils.DisplayInKompas(kompas, entry.Value, outputDirName + "\\" + fileName.Split('\\').Last(), styleFileName, table, 1);
                Utils.DisplayInKompas(kompas, entry.Value, outputDirName + "\\" + fileName.Split('\\').Last()+" (перечень)", styleFileName, table, 103);
            }
        }

        private async void btn_ok_Click(object sender, EventArgs e)
        {
            btn_ok.Enabled = false;

            try
            {
                Dictionary<string, string> columnNames = getColumnNames();
                bool val = Utils.createDirectory(outputDirName);
                if (val)
                {
                    btn_ok.Enabled = true;
                    return;
                }

                bool connect = cb_connect_all.Checked;

                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    var item = listBox1.Items[i].ToString();
                    int count = Convert.ToInt32(item.Split(' ').Last());
                    string newItem = Utils.changeFileCount(item, count * (int)multiplier.Value);
                    listBox1.Items[i] = newItem;
                }

                await Task.Run(() => runExcelToKompas(connect, columnNames));
                info_label.Text = "Готово!";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка в процессе конвертации!\n\nТекст ошибки: " + ex.Message);
                info_label.Text = "Ошибка в процессе выполнения! Перезагрузите ПО!";
            }

            btn_ok.Enabled = true;

            try { kompas.Quit(); } catch (Exception ex) { }
            try { excelApp.Quit(); } catch (Exception ex) { }
            InitializeApplications(false);
        }

        private Dictionary<string, string> getColumnNames()
        {
            Dictionary<string, string> output = new Dictionary<string, string>();
            if (format_name.Text != "") output["Формат"] = format_name.Text;
            if (zone_name.Text != "") output["Зона"] = zone_name.Text;
            if (position_name.Text != "") output["Позиция"] = position_name.Text;
            if (mark_name.Text != "") output["Обозначение"] = mark_name.Text;
            if (name_name.Text != "") output["Наименование"] = name_name.Text;
            if (count_name.Text != "") output["Количество"] = count_name.Text;
            if (add_name.Text != "") output["Примечание"] = add_name.Text;
            if (mass_name.Text != "") output["Масса"] = mass_name.Text;
            if (material_name.Text != "") output["Материал"] = material_name.Text;
            if (user_name.Text != "") output["Пользователь"] = user_name.Text;
            if (code_name.Text != "") output["Код"] = code_name.Text;
            if (manufacturer_name.Text != "") output["Изготовитель"] = manufacturer_name.Text;
            if (subsection_name.Text != "") output["Тип изделия"] = subsection_name.Text;
            return output;
        }

        private void WriteTextSafe(string text)
        {
            if (info_label.InvokeRequired)
                info_label.Invoke(new SafeCallDelegate(WriteTextSafe), new object[] { text });    
            else
                info_label.Text = text;
        }

        private void btn_upd_Click(object sender, EventArgs e)
        {
            try
            {
                reload();
            } catch(Exception ex)
            {
                info_label.Text = "Ошибка загрузки файлов из папки с исходниками";
            }
            
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            if(inputDirName != null && dir_in.Text != "")
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    try
                    {
                        File.Copy(file, inputDirName + "\\" + Path.GetFileName(file), true);
                    } catch (Exception ex)
                    {
                        info_label.Text = "Ошибка! " + Path.GetFileName(file) + " не был добавлен.";
                    }


                    listBox1.Items.Add(file+", 1");
                }
                    
            }
        }
    }
}
