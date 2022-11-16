using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TranMinhDuc_201210096
{
    public partial class Form1 : Form
    {
        DBConfig db = new DBConfig();
        string nameTable1 = "QLXe";
        string khoaBangTable1 = "";
        public Form1()
        {
            InitializeComponent();
        }

        bool ExistsInDatabase()
        {
            if (int.Parse(db.GetValue($"SELECT COUNT(*) FROM {nameTable1} WHERE {khoaBangTable1} = '{textBox1.Text}'").ToString()) != 0)
            {
                MessageBox.Show("Bản ghi đã tồn tại");
                textBox1.Focus();
                return true;
            }
            return false;
        }

        private bool isCheck()
        {
            if (textBox1.Text.Trim() == "")
            {
               MessageBox.Show("Xin mời nhập ");
                return false;
            }
            if (textBox2.Text.Trim() == "")
            {
                MessageBox.Show("Xin mời nhập ");
                return false;
            }
            if (textBox3.Text.Trim() == "")
            {
                MessageBox.Show("Xin mời nhập ");
                return false;
            }
            if (textBox4.Text.Trim() == "")
            {
                MessageBox.Show("Xin mời nhập ");
                return false;
            }
            if (textBox5.Text.Trim() == "")
            {
                MessageBox.Show("Xin mời nhập ");
                return false;
            }
            if (textBox6.Text.Trim() == "")
            {
                MessageBox.Show("Xin mời nhập ");
                return false;
            }
            if (picAnh.ImageLocation == "")
            {
                MessageBox.Show("Xin mời chọn ảnh");
                return false;
            }
            return true;
        }

        private void CleanInput()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            picAnh.ImageLocation = "";
            picAnh.Image = null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //"Select * from QLXe"
            dataGridView1.DataSource = db.table("");

            //DataTable dt = db.table("Select * from QLXiLanh");
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //  txtXiLanh.Items.Add(dt.Rows[i][1].ToString());
            //}
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            if (isCheck() || !ExistsInDatabase())
            {
                string query = $"insert into {nameTable1} values ('{textBox1.Text}',N'{textBox2.Text}',N'{textBox3.Text}','{textBox4.Text}','{textBox5.Text}','{textBox6.Text}','{picAnh.ImageLocation}')";
                try
                {
                    if (MessageBox.Show("Bạn có muốn thêm vào không ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        db.Excute(query);
                        Form1_Load(sender, e);
                        CleanInput();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Không thêm được !");
                }
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (isCheck())
            {
                try
                {
                    if (MessageBox.Show("Bạn có muốn sửa không ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        db.Excute($"update {nameTable1} set SoMay = '{textBox3.Text}', MaMau = '{textBox4.Text}' where {khoaBangTable1} = '{textBox1.Text}'");

                        Form1_Load(sender, e);
                        CleanInput();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Sửa lỗi !");
                }
            }
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Không được để trống ... !");
                return;
            }
            try
                {
                    if (MessageBox.Show("Bạn có muốn xóa không ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        db.Excute($"delete from {nameTable1} where {khoaBangTable1} = '{textBox1.Text}'");
                        Form1_Load(sender, e);
                        CleanInput();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Không tìm thấy ... để xóa!");
                }
        }
        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            CleanInput();
            dataGridView1.DataSource = db.table("");
        }

        private void btnTim_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Bạn có muốn tìm kiếm không ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    if (textBox2.Text == "" && textBox3.Text == "")
                    {
                        MessageBox.Show("Hãng xe và mã màu không được để trống !");
                        return;
                    }

                    // điều kiện tìm truy vấn
                    DataTable dataTable = db.table($"Select * from {nameTable1} where HangXe = '{textBox2.Text}' and MaMau = {textBox4.Text}");
                    if (dataTable.Rows.Count == 0)
                    {
                        MessageBox.Show("Không tìm thấy !");
                    }
                    else
                    {
                        dataGridView1.DataSource = dataTable;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Tìm lỗi !");
            }
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc muốn thoát không?", "Error", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textBox1.Text = dataGridView1[0, e.RowIndex].Value.ToString();
                textBox2.Text = dataGridView1[1, e.RowIndex].Value.ToString();
                textBox3.Text = dataGridView1[2, e.RowIndex].Value.ToString();
                textBox4.Text = dataGridView1[3, e.RowIndex].Value.ToString();
                textBox5.Text = dataGridView1[4, e.RowIndex].Value.ToString();
                textBox6.Text = dataGridView1[5, e.RowIndex].Value.ToString();
                picAnh.ImageLocation = dataGridView1[6, e.RowIndex].Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void btnAnh_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "JPEG Files (*.jpeg)|*.jpeg|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif|All Files|*.*";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                picAnh.ImageLocation = openFile.FileName;
            }
        }

        private void btnXuat_Click(object sender, EventArgs e)
        {
            if (txtHangXe.Text.Trim() == "")
            {
                MessageBox.Show("Chưa nhập hãng");
                return;
            }
            Excel.Application exApp = new Excel.Application();
            Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
            // Không insert cell 4 vì đề bài đang yêu cầu tim hãng xe lên t sẽ ko cần in hãng xe ra (cell4 là hãng xe)
            exSheet.get_Range("B2").Font.Bold = true;
            exSheet.get_Range("B2").Value = "DANH SÁCH XE " + txtHangXe.Text;
            exSheet.get_Range("A3").Value = "Số khung";
            exSheet.get_Range("B3").Value = "Số máy";
            exSheet.get_Range("C3").Value = "Mã màu";
            exSheet.get_Range("D3").Value = "Hãng xe";
            exSheet.get_Range("E3").Value = "Dung tích xi lanh";
            exSheet.get_Range("F3").Value = "Tên xe";
            exSheet.get_Range("G3").Value = "Ảnh";

            // thay đổi phụ thuộc bài toán
            int n = dataGridView1.Rows.Count - 1;
            int index = 0;
            for (int i = 0; i < n; i++)
            {
                string tenHang = dataGridView1.Rows[i].Cells[4].Value.ToString();
                if (tenHang.Equals(txtHangXe.Text.Trim()))
                {
                    exSheet.get_Range("A" + (index + 4).ToString()).Value = dataGridView1.Rows[i].Cells[0].Value;
                    exSheet.get_Range("B" + (index + 4).ToString()).Value = dataGridView1.Rows[i].Cells[1].Value;
                    exSheet.get_Range("C" + (index + 4).ToString()).Value = dataGridView1.Rows[i].Cells[2].Value;
                    exSheet.get_Range("D" + (index + 4).ToString()).Value = dataGridView1.Rows[i].Cells[3].Value;
                    exSheet.get_Range("E" + (index + 4).ToString()).Value = dataGridView1.Rows[i].Cells[4].Value;
                    exSheet.get_Range("F" + (index + 4).ToString()).Value = dataGridView1.Rows[i].Cells[5].Value;
                    exSheet.get_Range("F" + (index + 4).ToString()).Value = dataGridView1.Rows[i].Cells[6].Value;
                    index++;
                }

            }

            //
            exBook.Activate();

            //
            SaveFileDialog sfD = new SaveFileDialog();
            sfD.Filter = "Excel Files (*.xlsx)|*.xlsx";
            sfD.ShowDialog();
            exBook.SaveAs(sfD.FileName);

            //
            exApp.Quit();
        }
    }
}
