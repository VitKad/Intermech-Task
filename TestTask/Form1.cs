using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace TestTask
{
    public partial class Form1 : Form
    {
        private SqlConnection intermechConnection = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void Request()
        {
            //Загрузка данных по написаному SQL запросу
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                @"SELECT 
                objects.F_OBJECT_ID as 'Номер объекта',
                types.F_OBJECT_TYPE as 'Идентификатор типа объекта',
                types.F_OBJ_NAME as 'Наименование объекта',
                views.CAPTION as 'Описание'
                FROM INTERMECH_BASE.dbo.IMS_OBJECTS as objects
                INNER JOIN INTERMECH_BASE.dbo.IMS_OBJECT_TYPES as types
                ON objects.F_OBJECT_TYPE = types.F_OBJECT_TYPE
                INNER JOIN INTERMECH_BASE.dbo.IMS_OBJECTS_VIEW as views
                ON objects.F_OBJECT_ID = views.F_OBJECT_ID", intermechConnection);

            DataSet dataSet = new DataSet();

            dataAdapter.Fill(dataSet);

            //загрузка данных в GridView
            dataGridView1.DataSource = dataSet.Tables[0];
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //подключение к базе данных
            intermechConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["intermech"].ConnectionString);

            intermechConnection.Open();

            //выполение запроса
            Request();

            //заполнение comboBox значениями из таблицы типа объектов
            comboBox1.Items.Add("Все объекты");

            SqlCommand cmd = new SqlCommand(@"SELECT 
                types.F_OBJ_NAME as 'Наименование объекта'
                FROM INTERMECH_BASE.dbo.IMS_OBJECT_TYPES as types", intermechConnection);

            SqlDataReader DR = cmd.ExecuteReader();

            while (DR.Read())
            {
                comboBox1.Items.Add(DR[0]);

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //выгрузка элементов таблицы в Excel файл
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;

            int i, j;
            for (i = 0; i <= dataGridView1.RowCount - 2; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    wsh.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText.ToString();
                    wsh.Cells[i+2, j+1] = dataGridView1[j, i].Value.ToString();
                }
            }

            wsh.Columns.EntireColumn.AutoFit();
            exApp.Visible = true;
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Request();
            //поиск в таблице значений удовлетворяющих textBox
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter =
                 String.Format("CONVERT([Номер объекта], 'System.String') LIKE '{0}%'", textBox1.Text);

        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Request();
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter =
              String.Format("CONVERT([Наименование объекта], 'System.String') LIKE '{0}'", comboBox1.SelectedItem);
            if (comboBox1.SelectedItem == "Все объекты")
            {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter =
             String.Format("CONVERT([Наименование объекта], 'System.String') LIKE '%'");
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            Request();
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter =
              String.Format("CONVERT([Описание], 'System.String') LIKE '%{0}%'", textBox3.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //если значение не выбрано, то не ывполнять запрос
            if (dataGridView1.CurrentCell == null)
            {
                return;
            }
           
            SqlDataAdapter structureAdapter = new SqlDataAdapter(
               String.Format(
               @"SELECT DISTINCT
                objects.F_OBJECT_ID as 'Номер объекта', 
                types.F_OBJ_NAME as 'Наименование объекта',
                attributes.F_NAME as 'Наименование атрибута',
                object_attributes.F_STRING_VALUE as 'Значение атрибута',
                relations_types.F_DESCRIPTION as 'Тип связи'
                FROM INTERMECH_BASE.dbo.IMS_OBJECTS as objects
                LEFT JOIN INTERMECH_BASE.dbo.IMS_OBJECT_TYPES as types
                ON objects.F_OBJECT_TYPE = types.F_OBJECT_TYPE
                LEFT JOIN INTERMECH_BASE.dbo.IMS_OBJECT_ATTRS as object_attributes
                ON objects.F_OBJECT_ID = object_attributes.F_OBJECT_ID
                LEFT JOIN INTERMECH_BASE.dbo.IMS_ATTRIBUTES as attributes
                ON object_attributes.F_ATTRIBUTE_ID = attributes.F_ATTRIBUTE_ID
                LEFT JOIN INTERMECH_BASE.dbo.IMS_RELATIONS as relations
                ON objects.F_ID = relations.F_PART_ID
                LEFT JOIN INTERMECH_BASE.dbo.IMS_RELATION_TYPES as relations_types
                ON relations.F_RELATION_TYPE = relations_types.F_RELATION_TYPE
                WHERE objects.F_ID IN (
                SELECT relations.F_PART_ID as 'Потомок'
                FROM INTERMECH_BASE.dbo.IMS_RELATIONS as relations
                WHERE F_PROJ_ID = {0}) AND (object_attributes.F_ATTRIBUTE_ID = 9 OR object_attributes.F_ATTRIBUTE_ID = 10 OR object_attributes.F_ATTRIBUTE_ID = 1000 OR object_attributes.F_ATTRIBUTE_ID IS NULL);
             
                ", dataGridView1.CurrentCell.Value.ToString()
            ), intermechConnection);

            DataSet structureSet = new DataSet();

            structureAdapter.Fill(structureSet);

            dataGridView1.DataSource = structureSet.Tables[0];
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Сброс
            Request();
        }
    }
}
