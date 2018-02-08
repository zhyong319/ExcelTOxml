using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.IO;
using System.Xml;

namespace ExcelTOxml
{
    public partial class Form1 : Form
    {
        string excelFilePath;
        public Form1()
        {
            InitializeComponent();
           
        }

        private void button1_Click(object sender, EventArgs e) //加载EXCEL文件
        {
            OpenFileDialog f = new OpenFileDialog();
            //f.Filter = "Excel files(*.xlsx) | *.xlsx |All files(*.*) | *.*";
            if (f.ShowDialog() != DialogResult.OK) return;
            excelFilePath = Path.GetFullPath(f.FileName);
            this.textBox1.Text = excelFilePath;
        }

        private void button2_Click(object sender, EventArgs e) //生成XML文件
        {
            if (this.textBox1.Text == "")
            {
                MessageBox.Show("请先加载需要转换的EXCEL文档！！！");
                return;
            }
            if (this.textBox2.Text == "")
            {
                MessageBox.Show("请指定生成的XML文件保存位置！！！");
                return;
            }


            string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意                                                                                                                                            //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; //Office 07以下版本 
            using (OleDbConnection connection = new OleDbConnection(connstring))
            {
                connection.Open();
                DataTable sheetsName = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
                string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
                string sqlcommand = string.Format("SELECT * FROM [{0}]", firstSheetName); //查询字符串
                OleDbCommand command = new OleDbCommand(sqlcommand, connection);               
                OleDbDataReader excelReader = command.ExecuteReader();
                string xmlfilePath = CreateNewXmlFile(); //创建XML文件

                while (excelReader.Read()) //读取全部数据
                {
                    // MessageBox.Show(excelReader.GetString(0) + ", " + excelReader.GetString(1));
                    try
                    {
                        WriteStrToXMLex(excelReader.GetString(0), excelReader.GetString(1), xmlfilePath);
                    }
                    catch (InvalidCastException)
                    {
                        MessageBox.Show("转换已中止！\n\n请将EXCEL单元格的数据数据类型更改为[文本]类型");
                        return;
                    }
                }

                // always call Close when done reading.
                excelReader.Close();
                MessageBox.Show("恭喜！转换操作完成!","结果...",MessageBoxButtons.OK);
            }
        }

        private void WriteStrToXMLex(string section, string name, string fileName)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(fileName);
            XmlElement root = doc.DocumentElement;
            XmlElement numberElemnet = doc.CreateElement("USER"); //创建root根下number子节点
            XmlElement childElemnet_partment = doc.CreateElement("Section");  //创建number子节点下的section节点
            XmlElement childElement_name = doc.CreateElement("Name");//创建number子节点下的name节点
            childElemnet_partment.InnerText = section; //设置子节点的文本
            childElement_name.InnerText = name;//设置子节点的文本
            numberElemnet.AppendChild(childElemnet_partment); //添加子节点
            numberElemnet.AppendChild(childElement_name);
            root.AppendChild(numberElemnet);

            doc.Save(fileName);
        }

        private void button3_Click(object sender, EventArgs e) //选择XML保存位置
        {
            
             FolderBrowserDialog fd = new FolderBrowserDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                this.textBox2.Text = fd.SelectedPath;
            }
            

        }

        private string CreateNewXmlFile() //创建新文件 
        {
            DateTime dt = DateTime.Now;
           
            string _newFilePath;
            string fileName = "\\UserInfoLib";
            string extName = ".xml";
            _newFilePath = this.textBox2.Text + fileName + dt.Hour.ToString() + dt.Minute.ToString() + dt.Second.ToString() + extName;
            XmlDocument xml = new XmlDocument();
            XmlDeclaration decl = xml.CreateXmlDeclaration("1.0", "utf-8", null);
            xml.AppendChild(decl); //创建声明
            XmlElement rootEle = xml.CreateElement("ROOT");
            xml.AppendChild(rootEle); //创建根节点
            xml.Save(_newFilePath);//保存文件
           return _newFilePath;
        }

        private void button4_Click(object sender, EventArgs e)  //HELP帮助文档
        {
            Form2 f = new Form2();
            
            f.ShowDialog();
            
        }
    }
}
