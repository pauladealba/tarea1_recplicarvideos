namespace videosclase
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Paso 0: Condiciones al vacio
            if (textBox1.Text.Equals("") ||
               textBox2.Text.Equals(""))
            {
                MessageBox.Show("Los numeros tiene que ser MAYOR que cero, NO VACIOS");
                return;

            }
            //Paso 1: Inicializacion de Parametros
            int totalValores = Convert.ToInt32(textBox1.Text);
            int valorMuestra = Convert.ToInt32(textBox2.Text);
            //Paso 2: Declarar clase algoritmo genetico
            AlgoritmoSimulacion algoritmo = new AlgoritmoSimulacion();
            //Paso 3: Llamar metodo principal
            List<int> listaEnteros = algoritmo.GenerarValores(totalValores);
            //Paso 4: Llenar el grid
            llenarGrid(listaEnteros);

        }
        public void llenarGrid(List<int> lista)
        {
            //Paso 0: indicas el numero de columnas
            string numeroColumna1 = "1";
            string numeroColumna2 = "2";

            //Paso 1: determinas la cantidad de columnas
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add(numeroColumna1, "Id");
            dataGridView1.Columns.Add(numeroColumna2, "Valor");

            //Paso 2: recorres el grid para cada fila llenas los valores aleatorios
            for (int i = 0; i < lista.Count; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[Int32.Parse(numeroColumna1) - 1].Value = (i + 1).ToString();
                dataGridView1.Rows[i].Cells[Int32.Parse(numeroColumna2) - 1].Value = lista[i].ToString();
            }
        }

        public void DescargaExcel(DataGridView data)
        {
            //Paso 0: instalar complemente de excel
            Microsoft.Office.Interop.Excel.Application exportarExcel = new Microsoft.Office.Interop.Excel.Application();
            exportarExcel.Application.Workbooks.Add(true);
            int indiceColumna = 0;
            //Paso 1: construye columna y los nombres de las cabeceras}
            foreach (DataGridViewColumn columna in data.Columns)
            {
                indiceColumna++;
                exportarExcel.Cells[1, indiceColumna] = columna.HeaderText;
            }
            //Paso2: construyes filas y llenas valores
            int indiceFilas = 0;
            foreach (DataGridViewRow fila in data.Rows)
            {
                indiceFilas++;
                indiceColumna = 0;
                foreach (DataGridViewColumn columna in data.Columns)
                {
                    indiceColumna++;
                    exportarExcel.Cells[indiceFilas + 1, indiceColumna] = fila.Cells[columna.Name].Value;
                }
            }
            //Paso 3: visibilidad
            exportarExcel.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DescargaExcel(dataGridView1);
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
