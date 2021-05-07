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
    public partial class Frm2021 : Form
    {
        public Frm2021()
        {
            InitializeComponent();
        }

        xls.Application a = new xls.Application();
        int i = 7;
        DialogBoxPorEspecialidad especial = new DialogBoxPorEspecialidad();
        DialogBoxPorSemestre seme = new DialogBoxPorSemestre();

        private void Frm2021_Load(object sender, EventArgs e)
        {
            a.Workbooks.Open(Application.StartupPath + @"\formato2021.xlsx");
            while (a.ActiveWorkbook.ActiveSheet.Cells(i, 1).Value != null)
            {
                i++;
            }
            //i--;
        }

        private void btnLeer_Click(object sender, EventArgs e)
        {
            lvCaracteristicas.Items.Clear();
            int x = 6;
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
                lvCaracteristicas.Items.Add(lista);
                x++;
            }
        }

        private void btnEscribir_Click(object sender, EventArgs e)
        {
            string matricula = txtMatricula.Text;
            string paterno = txtPaterno.Text;
            string materno = txtMaterno.Text;
            string nombre = txtNombre.Text;
            string especialidad = cbEspecialidad.Text;
            string semestre = cbSemestre.Text;
            string servicio = cbServicio.Text;
            string practicas = cbPracticas.Text;
            string residencias = cbResidencias.Text;
            string certificaciones = cbCertificaciones.Text;
            string toefl = cbToefl.Text;
            txtMatricula.Clear();
            txtPaterno.Clear();
            txtMaterno.Clear();
            txtNombre.Clear();
            cbEspecialidad.ResetText();
            cbSemestre.ResetText();
            cbServicio.ResetText();
            cbPracticas.ResetText();
            cbResidencias.ResetText();
            cbCertificaciones.ResetText();
            cbToefl.ResetText();

            a.ActiveWorkbook.Worksheets[1].Cells(i, 1).Value = i - 5;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 2).Value = matricula;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 3).Value = paterno;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 4).Value = materno;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 5).Value = nombre;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 6).Value = especialidad;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 7).Value = semestre;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 8).Value = servicio;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 9).Value = practicas;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 10).Value = residencias;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 11).Value = certificaciones;
            a.ActiveWorkbook.ActiveSheet.Cells(i, 12).Value = toefl;
            i++;
            a.ActiveWorkbook.Save();
            MessageBox.Show("Se agregaron los datos al excel", "Lectura y Escritura", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            lvCaracteristicas.Items.Clear();
        }

        private void btnBEspecialidad_Click(object sender, EventArgs e)
        {
            especial.Show();
        }

        private void btnBSemestre_Click(object sender, EventArgs e)
        {
            seme.Show();
        }

        private void btnPEspecialidad_Click(object sender, EventArgs e)
        {
            int x = 6;
            int y = 6;
            int info = 0;
            int meca = 0;
            int gest = 0;
            int elec = 0;
            int indu = 0;
            int reno = 0;
            int total = 0;
            while (a.ActiveWorkbook.ActiveSheet.Cells(y, 1).Value != null)
            {
                total++;
                y++;
            }
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value != null)
            {
                string i = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();

                switch (i)
                {
                    case "Informática":
                        info++;
                        break;
                    case "Mecánica":
                        meca++;
                        break;
                    case "Gestión Empresarial":
                        gest++;
                        break;
                    case "Electrónica":
                        elec++;
                        break;
                    case "Industrial":
                        indu++;
                        break;
                    case "Energías Renovables":
                        reno++;
                        break;
                }
                x++;
            }

            decimal tInfo = (info * 100) / total;
            decimal tMeca = (meca * 100) / total;
            decimal tGest = (gest * 100) / total;
            decimal tElec = (elec * 100) / total;
            decimal tIndu = (indu * 100) / total;
            decimal tReno = (reno * 100) / total;

            MessageBox.Show("Porcentaje por especialidad: " +
                "Informática: " + tInfo + "%   " +
                "Mecánica: " + tMeca + "%   " +
                "Gestión Empresarial: " + tGest + "%   " +
                "Electrónica: " + tElec + "%   " +
              "Industial: " + tIndu + "%   " +
              "Energías Renovables: " + tReno + "%   ");
        }

        private void btnPSemestre_Click(object sender, EventArgs e)
        {
            int x = 6;
            int y = 6;
            int primero = 0;
            int segundo = 0;
            int tercero = 0;
            int cuarto = 0;
            int quinto = 0;
            int sexto = 0;
            int septimo = 0;
            int octavo = 0;
            int noveno = 0;
            int decimo = 0;
            int onceavo = 0;
            int doceavo = 0;
            int total = 0;
            while (a.ActiveWorkbook.ActiveSheet.Cells(y, 1).Value != null)
            {
                total++;
                y++;
            }
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value != null)
            {
                string b = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();

                switch (b)
                {
                    case "1":
                        primero++;
                        break;
                    case "2":
                        segundo++;
                        break;
                    case "3":
                        tercero++;
                        break;
                    case "4":
                        cuarto++;
                        break;
                    case "5":
                        quinto++;
                        break;
                    case "6":
                        sexto++;
                        break;
                    case "7":
                        septimo++;
                        break;
                    case "8":
                        octavo++;
                        break;
                    case "9":
                        noveno++;
                        break;
                    case "10":
                        decimo++;
                        break;
                    case "11":
                        onceavo++;
                        break;
                    case "12":
                        doceavo++;
                        break;
                }
                x++;
            }
            decimal t1 = (primero * 100) / total;
            decimal t2 = (segundo * 100) / total;
            decimal t3 = (tercero * 100) / total;
            decimal t4 = (cuarto * 100) / total;
            decimal t5 = (quinto * 100) / total;
            decimal t6 = (sexto * 100) / total;
            decimal t7 = (septimo * 100) / total;
            decimal t8 = (octavo * 100) / total;
            decimal t9 = (noveno * 100) / total;
            decimal t10 = (decimo * 100) / total;
            decimal t11 = (onceavo * 100) / total;
            decimal t12 = (doceavo * 100) / total;

            MessageBox.Show("Porcentaje por semestre: " +
              "1ro: " + t1 + "%   " +
              "2do: " + t2 + "%   " +
              "2ro: " + t3 + "%   " +
              "4to: " + t4 + "%   " +
              "5to: " + t5 + "%   " +
              "6to: " + t6 + "%   " +
              "7mo: " + t7 + "%   " +
              "8vo: " + t8 + "%   " +
              "9no:" + t9 + "%   " +
              "10mo:" + t10 + "%   " +
              "11vo:" + t11 + "%   " + "12vo:" + t12 + "%   ");
        }

        private void btnPServicio_Click(object sender, EventArgs e)
        {
            int x = 6;
            int y = 6;
            int si = 0;
            int no = 0;
            int total = 0;
            while (a.ActiveWorkbook.ActiveSheet.Cells(y, 1).Value != null)
            {
                total++;
                y++;
            }
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value != null)
            {
                string c = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();

                switch (c)
                {
                    case "Sí":
                        si++;
                        break;
                    case "No":
                        no++;
                        break;
                }
                x++;
            }
            decimal tSi = (si * 100) / total;
            decimal tNo = (no * 100) / total;

            MessageBox.Show("Porcentaje por servicio social: " +
              "Sí: " + tSi + "%   " +
              "No:" + tNo + "%   ");
        }

        private void btnPProfesionales_Click(object sender, EventArgs e)
        {
            int x = 6;
            int y = 6;
            int si = 0;
            int no = 0;
            int total = 0;
            while (a.ActiveWorkbook.ActiveSheet.Cells(y, 1).Value != null)
            {
                total++;
                y++;
            }
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value != null)
            {
                string d = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();

                switch (d)
                {
                    case "Sí":
                        si++;
                        break;
                    case "No":
                        no++;
                        break;
                }
                x++;
            }
            decimal tSi = (si * 100) / total;
            decimal tNo = (no * 100) / total;

            MessageBox.Show("Porcentaje por prácticas profesionales: " +
              "Sí: " + tSi + "%   " +
              "No:" + tNo + "%   ");
        }

        private void btnPResidenciales_Click(object sender, EventArgs e)
        {
            int x = 6;
            int y = 6;
            int si = 0;
            int no = 0;
            int total = 0;
            while (a.ActiveWorkbook.ActiveSheet.Cells(y, 1).Value != null)
            {
                total++;
                y++;
            }
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value != null)
            {
                string f = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();

                switch (f)
                {
                    case "Sí":
                        si++;
                        break;
                    case "No":
                        no++;
                        break;
                }
                x++;
            }
            decimal tSi = (si * 100) / total;
            decimal tNo = (no * 100) / total;

            MessageBox.Show("Porcentaje por prácticas residenciales: " +
              "Sí: " + tSi + "%   " +
              "No:" + tNo + "%   ");
        }

        private void btnCertificacion_Click(object sender, EventArgs e)
        {
            int x = 6;
            int y = 6;
            int si = 0;
            int no = 0;
            int total = 0;
            while (a.ActiveWorkbook.ActiveSheet.Cells(y, 1).Value != null)
            {
                total++;
                y++;
            }
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value != null)
            {
                string g = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();

                switch (g)
                {
                    case "Sí":
                        si++;
                        break;
                    case "No":
                        no++;
                        break;
                }
                x++;
            }
            decimal tSi = (si * 100) / total;
            decimal tNo = (no * 100) / total;

            MessageBox.Show("Porcentaje por certificaciones: " +
              "Sí: " + tSi + "%   " +
              "No:" + tNo + "%   ");
        }

        private void btnPToefl_Click(object sender, EventArgs e)
        {
            int x = 6;
            int y = 6;
            int si = 0;
            int no = 0;
            int total = 0;
            while (a.ActiveWorkbook.ActiveSheet.Cells(y, 1).Value != null)
            {
                total++;
                y++;
            }
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value != null)
            {
                string h = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();

                switch (h)
                {
                    case "Sí":
                        si++;
                        break;
                    case "No":
                        no++;
                        break;
                }
                x++;
            }
            decimal tSi = (si * 100) / total;
            decimal tNo = (no * 100) / total;

            MessageBox.Show("Porcentaje por TOEFL: " +
              "Sí: " + tSi + "%   " +
              "No:" + tNo + "%   ");
        }

        private void Frm2021_FormClosed(object sender, FormClosedEventArgs e)
        {
            a.ActiveWorkbook.Close();
        }
    }
}
