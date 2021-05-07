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
    public partial class DialogBoxPorEspecialidad : Form
    {
        public DialogBoxPorEspecialidad()
        {
            InitializeComponent();
        }

        xls.Application a = new xls.Application();
        int i = 7;

        private void DialogBoxPorEspecialidad_Load(object sender, EventArgs e)
        {
            a.Workbooks.Open(Application.StartupPath + @"\formato2021.xlsx");
            while (a.ActiveWorkbook.ActiveSheet.Cells(i, 1).Value != null)
            {
                i++;
            }
            i--;
        }

        private void btnInformatica_Click(object sender, EventArgs e)
        {
            lvEspecialidad.Items.Clear();
            int x = 6;
            string i = "Informática";
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
                if (especialidad == i)
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
                    lvEspecialidad.Items.Add(lista);
                }
                x++;
            }
        }

        private void btnMecanica_Click(object sender, EventArgs e)
        {
            lvEspecialidad.Items.Clear();
            int x = 6;
            string i = "Mecánica";
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
                if (especialidad == i)
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
                    lvEspecialidad.Items.Add(lista);
                }
                x++;
            }
        }

        private void btnGestion_Click(object sender, EventArgs e)
        {
            lvEspecialidad.Items.Clear();
            int x = 6;
            string i = "Gestión Empresarial";
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
                if (especialidad == i)
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
                    lvEspecialidad.Items.Add(lista);
                }
                x++;

            }
        }

        private void btnElectronica_Click(object sender, EventArgs e)
        {
            lvEspecialidad.Items.Clear();
            int x = 6;
            string i = "Electrónica";
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
                if (especialidad == i)
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
                    lvEspecialidad.Items.Add(lista);
                }
                x++;
            }
        }

        private void btnIndustrial_Click(object sender, EventArgs e)
        {
            lvEspecialidad.Items.Clear();
            int x = 6;
            string i = "Industrial";
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
                if (especialidad == i)
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
                    lvEspecialidad.Items.Add(lista);
                }
                x++;
            }
        }

        private void btnEnergias_Click(object sender, EventArgs e)
        {
            lvEspecialidad.Items.Clear();
            int x = 6;
            string i = "Energías Renovables";
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
                if (especialidad == i)
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
                    lvEspecialidad.Items.Add(lista);
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
