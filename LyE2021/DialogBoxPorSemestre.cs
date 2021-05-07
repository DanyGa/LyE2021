using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using xls = Microsoft.Office.Interop.Excel;

namespace LyE2021
{
    public partial class DialogBoxPorSemestre : Form
    {
        public DialogBoxPorSemestre()
        {
            InitializeComponent();
        }

        xls.Application a = new xls.Application();
        int i = 7;

        private void DialogBoxPorSemestre_Load(object sender, EventArgs e)
        {
            a.Workbooks.Open(Application.StartupPath + @"\formato2021.xlsx");
            while (a.ActiveWorkbook.ActiveSheet.Cells(i, 1).Value != null)
            {
                i++;
            }
            i--;
        }

        private void btn1_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string p = "1";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == p)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn2_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string s = "2";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == s)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn3_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string t = "3";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == t)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn4_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string c = "4";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == c)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn5_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string q = "5";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == q)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn6_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string se = "6";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == se)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn7_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string sep = "7";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == sep)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn8_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string o = "8";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == o)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn9_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string n = "9";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == n)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn10_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string d = "10";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == d)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn11_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string on = "11";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == on)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btn12_Click(object sender, EventArgs e)
        {
            lvSemestre.Items.Clear();
            int x = 6;
            string doce = "12";
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                string num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                string matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                string paterno = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                string materno = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                string nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                string especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                string semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                string servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                string practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                string residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                string certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                string toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                ListViewItem lista = new ListViewItem(num);
                if (semestre == doce)
                {
                    lista.SubItems.Add(matricula);
                    lista.SubItems.Add(paterno);
                    lista.SubItems.Add(materno);
                    lista.SubItems.Add(nombre);
                    lista.SubItems.Add(especialidad);
                    lista.SubItems.Add(semestre);
                    lista.SubItems.Add(servicio);
                    lista.SubItems.Add(practicas);
                    lista.SubItems.Add(residencias);
                    lista.SubItems.Add(certificaciones);
                    lista.SubItems.Add(toefl);
                    lvSemestre.Items.Add(lista);
                }
                x++;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Hide();
        }
    }
}
