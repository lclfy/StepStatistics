using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CCWin;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace StepStatistics
{


    public partial class StepStatistics : Skin_Mac
    {

        List<DataUnit> AllDatas;
        IWorkbook WorkBook;
        string TextData;
        string date;
        string FileName;
        public StepStatistics()
        {
            InitializeComponent();
        }

        private bool ImportFiles(int type)
        {
            //0为excel 1为图片
            if(type == 0)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();   //显示选择文件对话框 
                openFileDialog1.Multiselect = false;
                openFileDialog1.Filter = "Excel 文件 |*.xlsx;*.xls";
                //openFileDialog1.InitialDirectory = Application.StartupPath + "\\时刻表\\";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;
                IWorkbook workBook;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    String fileNames = "已选择：";
                    string fileName = openFileDialog1.FileName;
                    {
                        try
                        {
                            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                            if (fileName.IndexOf(".xlsx") > 0) // 2007版本  
                            {
                                try
                                {
                                    workBook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook  
                                    WorkBook = workBook;
                                    FileName = fileName;
                                    if (TextData.Length != 0)
                                    {
                                        button2.Enabled = true;
                                    }
                                }
                                catch (Exception e)
                                {
                                    MessageBox.Show("读取表格时出现错误\n" + fileName + "\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                }

                            }
                            else if (fileName.IndexOf(".xls") > 0) // 2003版本  
                            {
                                // try
                                {
                                    workBook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
                                    WorkBook = workBook;
                                    FileName = fileName;
                                    if (TextData.Length != 0)
                                    {
                                        button2.Enabled = true;
                                    }
                                }
                                // catch (Exception e)
                                {
                                    //  MessageBox.Show("读取问题查询表时出现错误\n" + fileName + "\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    // return false;
                                }
                            }
                            fileStream.Close();
                            fileStream.Close();
                        }
                        catch (IOException)
                        {
                             MessageBox.Show("读取表格时出现错误\n" + fileName, "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    fileNames = fileNames + fileName;
                    label4.Text = fileNames;
                }
            }
            else if(type == 1)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();   //显示选择文件对话框 
                openFileDialog1.Multiselect = false;
                openFileDialog1.Filter = "图片 |*.jpg;*.png";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string file = openFileDialog1.FileName;
                }


            }
            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AllDatas = new List<DataUnit>();
            TextData = "";
            date = "";
            button2.Enabled = false;
            FileName = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ImportFiles(0);
        }

        private void getData()
        {
            removeUnuseableWords();
            List<DataUnit> _allDU = new List<DataUnit>();
            string[] splitedText = TextData.Split('\n');
            int allCounter = 0;
            int nameCounter = 0;
            int stepCounter = 0;
            for(int i = 0; i < splitedText.Length; i++)
            {
                string _tempText = splitedText[i].Trim().Replace(" ","").Replace(" ","");
                //如果是空的就跳过
                if(_tempText.Trim().Length == 0)
                {
                    continue;
                }
                //分三种情况，只数字，只汉字，汉字数字叠加
                //遇到哪一种，总数+1，该种类计数+1，填入既定位置
                //汉字数字，直接添加一整个
                if(Regex.IsMatch(_tempText, @"^[\u4E00-\u9FA50-9]+$") &&
                    !Regex.IsMatch(_tempText, @"^[\u4E00-\u9FA5]*$")&&
                    !Regex.IsMatch(_tempText, @"^[0-9]*$"))
                {
                    //取汉字部分
                    string _name = Regex.Replace(_tempText, @"[0-9]", ""); //只留汉字
                    //取数字部分
                    string _step = Regex.Replace(_tempText, @"[^0-9]", ""); //
                    DataUnit _du = new DataUnit(_name, _step);
                    _allDU.Add(_du);
                    allCounter++;
                    nameCounter++;
                    stepCounter++;
                }
                //汉字，添加姓名部分
                else if (Regex.IsMatch(_tempText, @"^[\u4E00-\u9FA5]*$"))
                {
                    string _name = _tempText;
                    //有多余的
                    if(nameCounter < stepCounter && nameCounter < allCounter)
                    {
                        _allDU[nameCounter].name = _name;
                    }
                    //没有多余的，添加一个新的
                    if(nameCounter == allCounter)
                    {
                        DataUnit _du = new DataUnit(_name, "");
                        _allDU.Add(_du);
                        allCounter++;
                    }
                    nameCounter++;
                }
                //数字，添加步数部分
                else if(Regex.IsMatch(_tempText, @"^[0-9]*$"))
                {
                    string _step = _tempText;
                    //有多余的
                    if (stepCounter < nameCounter && stepCounter < allCounter)
                    {
                        _allDU[stepCounter].steps = _step;
                    }
                    else if (stepCounter == allCounter)
                    {
                        DataUnit _du = new DataUnit("", _step);
                        _allDU.Add(_du);
                        allCounter++;
                    }
                    stepCounter++;

                }
            }
            AllDatas = _allDU;
            string _showntext = "共计"+_allDU.Count+"人：";
            foreach (DataUnit _tempdu in _allDU)
            {
                _showntext = _showntext + "\n" + _tempdu.name + "-" + _tempdu.steps;
            }
            richTextBox2.Text = _showntext;
        }

        private void removeUnuseableWords()
        {
            TextData = TextData.Replace("郑容段列车长", "");
            TextData = TextData.Replace("郑客段列车长王凯：1", "王凯男");
            TextData = TextData.Replace("王蕊（京焦一）", "王蕊京焦");
            TextData = TextData.Replace("郑客段列车长", "");
            TextData = TextData.Replace("郑客段列车长士解元", "王静亮");
            TextData = TextData.Replace("郑客", "");
            TextData = TextData.Replace(".", "");
            TextData = TextData.Replace("-", "");
            TextData = TextData.Replace("…", "");
            TextData = TextData.Replace("A平安银行党志诚17", "党志诚").Replace("A平安银行党志诚", "党志诚");
            TextData = TextData.Replace("千家惠", "于家惠");
            TextData = TextData.Replace("134628", "");
            TextData = TextData.Replace("彭彭晓梦", "彭晓梦");
            TextData = TextData.Replace("张大鹏1063", "张大鹏");
            TextData = TextData.Replace("进步神速", "");
            TextData = TextData.Replace("迸步神速", "");
            TextData = TextData.Replace("今日冠军", "");
            TextData = TextData.Replace("棼", "梦").Replace("婪","梦");
            TextData = TextData.Replace("鵾", "鹍").Replace("鹖", "鹍").Replace("陶鹃", "陶鹍");
            TextData = TextData.Replace("吳", "吴");
            TextData = TextData.Replace("雃", "雅");
            TextData = TextData.Replace("崔崔婕", "崔婕");
            TextData = TextData.Replace("张张雪姣", "张雪姣");
            TextData = TextData.Replace("韩希頁", "韩希真");
            TextData = TextData.Replace("乇", "毛");
            TextData = TextData.Replace("王漬君", "王清君");
            TextData = TextData.Replace("千慧杰", "王慧杰");
            
            TextData = TextData.Replace("盂鑫", "孟鑫");
            TextData = TextData.Replace("兩", "雨");
           TextData = TextData.Replace("𡥄", "孟");
            TextData = TextData.Replace("𡝭", "娟");
            TextData = TextData.Replace("•", "");
             TextData = TextData.Replace("𣇈", "晓");
            TextData = TextData.Replace("𡈼", "王").Replace("壬","王"); 
             TextData = TextData.Replace("𦍋", "毕");
            TextData = TextData.Replace("高铁", "");
            TextData = TextData.Replace("一队", "");
            TextData = TextData.Replace("京武", "");
            TextData = TextData.Replace("京郑", "");
            TextData = TextData.Replace("呼和", "");
            TextData = TextData.Replace("京商", "");
            TextData = TextData.Replace("京银", "");
            TextData = TextData.Replace("京南", "");
            TextData = TextData.Replace("京西", "");
            TextData = TextData.Replace("一组", "");
            TextData = TextData.Replace("1组", "");
            TextData = TextData.Replace("二组", "");
            TextData = TextData.Replace("2组", "");
            TextData = TextData.Replace("三组", "");
            TextData = TextData.Replace("3组", "");
            TextData = TextData.Replace("四组", "");
            TextData = TextData.Replace("4组", "");
            TextData = TextData.Replace("五组", "");
            TextData = TextData.Replace("5组", "");
            TextData = TextData.Replace("六组", "");
            TextData = TextData.Replace("6组", "");
            TextData = TextData.Replace("悅", "悦");
            TextData = TextData.Replace("雲", "云");

            //去手机号
            TextData = Regex.Replace(TextData, @"\d{11}", "");

        }

        //写入excel文件
        private void writeData()
        {
            IWorkbook workbook = WorkBook;
            ISheet sheet = workbook.GetSheetAt(0);
            int dateColumn = 0;
            //在第一行找日期
            if(sheet.GetRow(0) == null || date.Length == 0)
            {
                MessageBox.Show("未填写日期或未在表格第一行找到日期");
                return;
            }
            IRow _rowdate = sheet.GetRow(0);
            bool hasGotDate = false;
            for(; dateColumn <= _rowdate.LastCellNum; dateColumn++)
            {
                if (_rowdate.GetCell(dateColumn) != null)
                {
                    if(_rowdate.GetCell(dateColumn).ToString().Length != 0)
                    {
                        if (_rowdate.GetCell(dateColumn).ToString().Trim().Equals(date))
                        {
                            //找到了，跳出
                            hasGotDate = true;
                            break;
                        }
                    }
                }
            }
            if (!hasGotDate)
            {
                MessageBox.Show("未填写日期或未在表格第一行找到日期");
                return;
            }
            //没找到的，在后面添加
            List<DataUnit> _hasNotFoundedDU = new List<DataUnit>();
            List<DataUnit> _allDU = AllDatas;
            //第一列是姓名，用姓名库在表里匹配一下，没有的话拎出来新建
            //表格里的所有名字
            List<DataUnit> _allNames = new List<DataUnit>();
            for (int i = 1;i<= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if(row == null)
                {
                    continue;
                }
                if(row.GetCell(0) != null)
                {
                    if(row.GetCell(0).ToString().Length != 0)
                        {
                        DataUnit _tempDU = new DataUnit(row.GetCell(0).ToString().Trim().Replace(" ", ""), "", i);
                            _allNames.Add(_tempDU);
                        }
                }
            }
            //匹配一下，没有的新建，有的写上
            foreach (DataUnit _du in AllDatas)
            {
                bool hasGotLineup = false;
                for (int _s = 0; _s < _allNames.Count; _s++)
                {
                    if (_du.name.Trim().Equals(_allNames[_s].name.Replace("（","").Replace("）","")))
                    {
                        if(sheet.GetRow(_allNames[_s].indexAt)== null)
                        {
                            sheet.CreateRow(_allNames[_s].indexAt);
                        }
                        IRow _tempRow = sheet.GetRow(_allNames[_s].indexAt);
                        if(_tempRow.GetCell(dateColumn) == null)
                        {
                            _tempRow.CreateCell(dateColumn);
                        }
                        _tempRow.GetCell(dateColumn).SetCellValue(_du.steps);
                        hasGotLineup = true;
                        break;
                    }
                }
                if(hasGotLineup == false)
                {
                    _hasNotFoundedDU.Add(_du);
                }
            }
            //没有的新建
            int lastRow = -1;
            //找最后一行(空的第一行)
            for (int ij=1;ij<sheet.LastRowNum;ij++)
            {
                if(sheet.GetRow(ij) == null)
                {
                    lastRow = ij;
                    break;
                }
                if(sheet.GetRow(ij).GetCell(0) == null)
                {
                    lastRow = ij;
                    break;
                }
                if (sheet.GetRow(ij).GetCell(0).ToString().Length == 0)
                {
                    lastRow = ij;
                    break;
                }
            }
            if(lastRow == -1)
            {
                lastRow = sheet.LastRowNum+1;
            }
            for (int counter = 0; counter < _hasNotFoundedDU.Count; counter++)
            {
                if(sheet.GetRow(lastRow + counter) == null)
                {
                    sheet.CreateRow(lastRow + counter);
                }
                IRow _tempRow = sheet.GetRow(lastRow + counter);
                if (_tempRow.GetCell(0) == null)
                {
                    _tempRow.CreateCell(0);
                }
                _tempRow.GetCell(0).SetCellValue(_hasNotFoundedDU[counter].name);
                if(_tempRow.GetCell(dateColumn) == null)
                {
                    _tempRow.CreateCell(dateColumn);
                }
                _tempRow.GetCell(dateColumn).SetCellValue(_hasNotFoundedDU[counter].steps);
            }
            /*重新修改文件指定单元格样式*/
            FileStream fs1 = File.OpenWrite(FileName);
            workbook.Write(fs1);
            fs1.Close();
            workbook.Close();
            MessageBox.Show("写入完成", "提示", MessageBoxButtons.OK);
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            TextData = richTextBox1.Text;
            if (TextData.Length != 0 && TextData != null)
            {
                getData();
                if(WorkBook != null)
                {
                    button2.Enabled = true;
                }
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            writeData();
            textBox1.Text = "X.X";
        }

        //选择图片
        private void button3_Click(object sender, EventArgs e)
        {
            ImportFiles(1);
        }

        //输入日期
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            date = textBox1.Text;
        }
    }

    public class DataUnit
    {
        public string name { get; set; }
        public string steps { get; set; }
        public int indexAt { get; set; }

        public DataUnit(string _name, string _steps,int _indexAt = -1)
        {
            name = _name;
            steps = _steps;
            indexAt = _indexAt;
        }

    }
}
