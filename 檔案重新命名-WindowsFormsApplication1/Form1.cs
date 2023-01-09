using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using io = System.IO;

namespace 檔案重新命名_WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string[] filesList;
        List< string>imagesList=new List<string>();
        public Form1()
        {
            InitializeComponent();
            SimpleList family = new SimpleList();

            // Populate the List            
            family.Add("沛榮");
            family.Add("玫儀");
            family.Add("靖F");
            family.Add("JT");
            family.Add("five");
            family.Add("six");
            family.Add("seven");
            family.Add("eight");
            family.PrintContents();
            comboBox1.DataSource = family;

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            //test the pull request!
            if (io.Directory.Exists(textBox1.Text) && io.Directory.Exists(textBox2.Text))
            {
                foreach (var itemF in io.Directory.GetFiles(textBox1.Text))
                {
                    io.FileInfo f = new io.FileInfo(itemF);
                    foreach (var itemT in io.Directory.GetFiles(textBox2.Text))
                    {
                        io.FileInfo ft = new io.FileInfo(itemT);
                        if (f.LastWriteTime == ft.LastWriteTime && f.Length == ft.Length && f.Name == ft.Name)
                        {

                            if (io.File.GetAttributes(itemF).ToString().IndexOf(io.FileAttributes.ReadOnly.ToString()) != -1)
                            {
                                io.File.SetAttributes(itemF, io.FileAttributes.Normal);
                            }
                            io.File.Delete(itemF);
                            break;
                        }
                    }
                }
                MessageBox.Show("done!");
            }
            else
            {
                MessageBox.Show("路徑有誤！");
            }
        }


        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox1.Text = Clipboard.GetText();
        }
        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox2.Text = Clipboard.GetText();
        }


        private void Form1_DragEnter(object sender, DragEventArgs e)
        {//https://wijtb.nctu.me/archives/269/
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;//调用DragDrop事件 
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {//https://www.google.com/search?rlz=1C1OKWM_zh-TWTW847TW847&ei=W7wBXbDjG8Kk8AWfkYGoBQ&q=c%23+%E6%8B%96%E6%8B%89%E7%89%A9%E4%BB%B6+%E8%B3%87%E6%96%99%E5%A4%BE&oq=c%23+%E6%8B%96%E6%8B%89%E7%89%A9%E4%BB%B6+%E8%B3%87%E6%96%99&gs_l=psy-ab.3.0.33i160l2.11461.20026..22176...0.0..0.99.232.3......0....1..gws-wiz.......0i30.KqdNNbPjI8A
            //https://wijtb.nctu.me/archives/269/
            string[] filePaths = (string[])e.Data.GetData(DataFormats.FileDrop);
            this.textBox1.Text = filePaths[0];
        }

        private void textBox2_DragDrop(object sender, DragEventArgs e)
        {
            string[] filePaths = (string[])e.Data.GetData(DataFormats.FileDrop);
            this.textBox2.Text = filePaths[0];
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                this.TopMost = true;
            }
            else
            {
                this.TopMost = false;
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void comboBox_Click(ComboBox cmbx)
        {
            string x = cmbx.SelectedValue.ToString();
            if (textBox2.Text == "重新命名預覽")
            {
                textBox2.Text = "";
            }
            textBox2.Text = textBox2.Text + x;

        }

        class SimpleList : IList //https://docs.microsoft.com/zh-tw/dotnet/api/system.collections.ilist?view=netframework-4.8
        {
            private object[] _contents = new object[8];
            private int _count;

            public SimpleList()
            {
                _count = 0;
            }

            // IList Members
            public int Add(object value)
            {
                if (_count < _contents.Length)
                {
                    _contents[_count] = value;
                    _count++;

                    return (_count - 1);
                }
                else
                {
                    return -1;
                }
            }

            public void Clear()
            {
                _count = 0;
            }

            public bool Contains(object value)
            {
                bool inList = false;
                for (int i = 0; i < Count; i++)
                {
                    if (_contents[i] == value)
                    {
                        inList = true;
                        break;
                    }
                }
                return inList;
            }

            public int IndexOf(object value)
            {
                int itemIndex = -1;
                for (int i = 0; i < Count; i++)
                {
                    if (_contents[i] == value)
                    {
                        itemIndex = i;
                        break;
                    }
                }
                return itemIndex;
            }

            public void Insert(int index, object value)
            {
                if ((_count + 1 <= _contents.Length) && (index < Count) && (index >= 0))
                {
                    _count++;

                    for (int i = Count - 1; i > index; i--)
                    {
                        _contents[i] = _contents[i - 1];
                    }
                    _contents[index] = value;
                }
            }

            public bool IsFixedSize
            {
                get
                {
                    return true;
                }
            }

            public bool IsReadOnly
            {
                get
                {
                    return false;
                }
            }

            public void Remove(object value)
            {
                RemoveAt(IndexOf(value));
            }

            public void RemoveAt(int index)
            {
                if ((index >= 0) && (index < Count))
                {
                    for (int i = index; i < Count - 1; i++)
                    {
                        _contents[i] = _contents[i + 1];
                    }
                    _count--;
                }
            }

            public object this[int index]
            {
                get
                {
                    return _contents[index];
                }
                set
                {
                    _contents[index] = value;
                }
            }

            // ICollection Members

            public void CopyTo(Array array, int index)
            {
                int j = index;
                for (int i = 0; i < Count; i++)
                {
                    array.SetValue(_contents[i], j);
                    j++;
                }
            }

            public int Count
            {
                get
                {
                    return _count;
                }
            }

            public bool IsSynchronized
            {
                get
                {
                    return false;
                }
            }

            // Return the current instance since the underlying store is not
            // publicly available.
            public object SyncRoot
            {
                get
                {
                    return this;
                }
            }

            // IEnumerable Members

            public IEnumerator GetEnumerator()
            {
                // Refer to the IEnumerator documentation for an example of
                // implementing an enumerator.
                throw new Exception("The method or operation is not implemented.");
            }

            public void PrintContents()
            {
                Console.WriteLine("List has a capacity of {0} and currently has {1} elements.", _contents.Length, _count);
                Console.Write("List contents:");
                for (int i = 0; i < Count; i++)
                {
                    Console.Write(" {0}", _contents[i]);
                }
                Console.WriteLine();
            }
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            //comboBox1.DroppedDown=true;

        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBox_Click(comboBox1);
        }

        private void comboBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Middle)
            {
                comboBox_Click(comboBox1);

            }

        }

        private void comboBox1_Click_1(object sender, EventArgs e)
        {
            comboBox1.DroppedDown = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")return;
            //if (MessageBox.Show("是否開始資料夾內的檔案重新命名？", "重新命名", MessageBoxButtons.OKCancel) == DialogResult.OK)
            //{
                string path = textBox1.Text;
                if (io.Directory.Exists(path))
                {
                    filesList= io.Directory.GetFiles(path);
                    
                    foreach (var item in filesList)
                    {
                        if ("jpg,png,bmp,gif,".IndexOf(item.Substring(item.Length - 3)) > -1)
                            imagesList.Add(item);                        
                    }
                }



                foreach (string fstr in filesList) //io.Directory.GetFiles(path))
                {
                    if (fstr == "") continue;
                    io.FileInfo f = new io.FileInfo(fstr);
                    //if (f.LastAccessTime< DateTime.Today)
                    //{
                    switch (f.Extension)
                    {
                        case ".jpg":
                            pictureBox1.ImageLocation = f.FullName;
                            // pictureBox1.Image = new Bitmap( f.FullName);//二式皆可
                            return;
                        //break;
                        case ".jpeg":
                        case ".png":
                        case ".bmp":
                        case ".tif":
                        case ".tiff":
                        default:
                            break;
                    }
                    //}

                }


            //}
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        int imageIndex=0;
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (imagesList.Count()== 0) return;
            //next image
            pictureBox1.ImageLocation = filesList[imageIndex++];

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        }

        private void pictureBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Middle)
            {

            }
        }
    }
}
