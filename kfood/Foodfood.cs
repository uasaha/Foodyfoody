using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.ML;
using KfoodML.Model;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace kfood
{
    public partial class Foodfood : Form
    {
        public Foodfood()
        {
            InitializeComponent();

            pbDragNDrop.AllowDrop = true;
        }

        private void PdDragNDropDragDrop(object sender, DragEventArgs e)
        {
            try
            {
                var directoryName = (string[])e.Data.GetData(DataFormats.FileDrop);
                pbDragNDrop.Image = Image.FromFile(directoryName[0], true);

                label3.Text = directoryName[0];
                label4.Visible = false;
            }

            catch (System.Exception a)
            {
                MessageBox.Show(a.Message);
            }
        }

        private void PbDragNDropDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? 
                DragDropEffects.Copy : DragDropEffects.None;
        }

        private void button1_Click(object sender, EventArgs e)
        {
                ModelInput PlusedImage = new ModelInput()
                {
                    ImageSource = label3.Text
                };
                var predictionResult = ConsumeModel.Predict(PlusedImage);

                label2.Visible = true;
                label2.Text = ($"{predictionResult.Prediction}");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            switch (label2.Text)
            {
                case "가지볶음":
                    richTextBox1.Text = "117\n9.3\n8.14\n2\n1.33\n2.63\n3.5";
                    break;
                case "간장게장":
                    richTextBox1.Text = "105\n1.89\n5.11\n16.8\n0.27\n2.12\n1.2";
                    break;
                case "갈비구이":
                    richTextBox1.Text = "240\n13.68\n0.89\n26.56\n5.15\n0.27\n0.1";
                    break;
                case "갈비찜":
                    richTextBox1.Text = "298\n21.79	\n7.72\n17.27\n8.44	\n4.7\n0.6";
                    break;
                case "갈치구이":
                    richTextBox1.Text = "\n96\n3.77\n3.51\n11.08\n0.87\n1.98\n0.5";
                    break;
                case "감자전":
                    richTextBox1.Text = "128\n7.44\n13.67\n1.74\n1.07\n1.15\n2.4";
                    break;
                case "감자조림":
                    richTextBox1.Text = "124\n4.8\n19.01\n2.25\n0.67\n2.11\n1.6";
                    break;
                case "감자채볶음":
                    richTextBox1.Text = "118\n6.03\n15.24\n1.63\n0.84\n1.74\n1.4";
                    break;
                case "감자탕":
                    richTextBox1.Text = "76\n2.79\n5.01\n7.37\n0.84\n0.49\n1";
                    break;
                case "갓김치":
                    richTextBox1.Text = "81\n1.75\n13.83\n5.89\n0.17\n2.56\n2.4";
                    break;
                case "건새우볶음":
                    richTextBox1.Text = "133\n3.15\n1.98\n23.26\n0.61\n0.12\n0.1";
                    break;
                case "경단":
                    richTextBox1.Text = "235\n1.37\n49.87\n4.78\n0.57\n4.59n2.4";
                    break;
                case "계란국":
                    richTextBox1.Text = "54\n2.95\n2.65\n4.32\n0.92\n1.12\n0.5";
                    break;
                case "계란말이":
                    richTextBox1.Text = "134\n8.96\n2.61\n10.18\n2.47\n1.12\n0.7";
                    break;
                case "계란찜":
                    richTextBox1.Text = "142\n10.11\n2.07\n10.05\n2.59\n0.95\n0.5";
                    break;
                case "계란후라이":
                    richTextBox1.Text = "194\n14.69\n0.93\n13.56\n4.09\n0.83\n0";
                    break;
                case "고등어구이":
                    richTextBox1.Text = "190\n12.6\n0.35\n17.59\n2.9\n0.09\n0";
                    break;
                case "고등어조림":
                    richTextBox1.Text = "143\n8.36\n4.13\n12.67\n1.95\n1.3\n1.2";
                    break;
                case "고사리나물":
                    richTextBox1.Text = "64\n3.39\n6.87\n3.24\n0.45\n1.78\n0.7";
                    break;
                case "고추장진미채볶음":
                    richTextBox1.Text = "278\n17.62\n22.53\n12.23\n3.05\n6.38\n5";
                    break;
                case "고추튀김":
                    richTextBox1.Text = "219\n	10.03\n	28.67\n	3.85\n	0.75\n	1.94\n	2.3";
                    break;
                case "곰탕_설렁탕":
                    richTextBox1.Text = "169\n	5.17\n	17.08\n	13.07\n	2.01\n	0.13\n	0.2";
                    break;
                case "곱창구이":
                    richTextBox1.Text = "134\n	9.24\n	1.19\n	11.15\n	1.98\n	0.04\n	0.1";
                    break;
                case "곱창전골":
                    richTextBox1.Text = "166\n	11.77\n	10.48\n	7.43\n	2.06\n	1.72\n	3.3";
                    break;
                case "과메기":
                    richTextBox1.Text = "177\n	9.95\n	0\n	20.25\n	2.25\n	0\n	0";
                    break;
                case "김밥":
                    richTextBox1.Text = "118\n	0.24\n	25.8\n	2.21\n	0.07\n	4.47\n	0.5";
                    break;
                case "김치볶음밥":
                    richTextBox1.Text = "149\n	4.12\n	25.03\n	2.92\n	0.62\n	0.82\n	1.3";
                    break;
                case "김치전":
                    richTextBox1.Text = "118\n	5.16\n	13.24\n	4.86\n	1.08\n	3.24\n	1.2";
                    break;
                case "김치찌개":
                    richTextBox1.Text = "60\n3.79\n3.23\n3.76\n0.95\n0.89\n0.7";
                    break;
                case "김치찜":
                    richTextBox1.Text = "131\n	0.9\n	28.21\n	5.6\n	0.2\n	3.42\n	2.9";
                    break;
                case "깍두기":
                    richTextBox1.Text = "32\n	0.59\n	6.41\n	1.24\n	0.11\n	2.89\n	2";
                    break;
                case "깻잎장아찌":
                    richTextBox1.Text = "74\n	1.56\n	12.88\n	6.03\n	0.18\n	2.12\n	2.7";
                    break;
                case "꼬막찜":
                    richTextBox1.Text = "81\n	1.06\n	2.8\n	13.94\n	0.1\n	0\n	0";
                    break;
                case "꽁치조림":
                    richTextBox1.Text = "134\n	8.75\n	4.29\n	9.89\n	2.03\n	1.68\n	1";
                    break;
                case "꽈리고추무침":
                    richTextBox1.Text = "68\n	4.08\n	8.21\n	2\n	0.59\n	2.83\n	3.3";
                    break;
                case "꿀떡":
                    richTextBox1.Text = "217\n	1.3\n	46.12\n	4.3\n	0.2\n	3.01\n	2";
                    break;
                case "나박김치":
                    richTextBox1.Text = "14\n	0.21\n	2.86\n	1.2\n	0.03\n	1.27\n	1.1";
                    break;
                case "누룽지":
                    richTextBox1.Text = "383\n	0.51\n	84.29\n	7.31\n	0.12\n	0\n	2.1";
                    break;
                case "닭갈비":
                    richTextBox1.Text = "195\n	8.99\n	14.39\n	13.4\n	2.42\n	1\n	1.1";
                    break;
                case "닭계장":
                    richTextBox1.Text = "95\n	4.85\n	2.44\n	10.46\n	1.25\n	0.28\n	0.4";
                    break;
                case "닭볶음탕":
                    richTextBox1.Text = "153\n	6.33\n	6.27\n	17.15\n	1.54\n	1.67\n	1";
                    break;
                case "더덕구이":
                    richTextBox1.Text = "82\n	2.65\n	15.16\n	3.27\n	0.44\n	0.86\n	6";
                    break;
                case "도라지무침":
                    richTextBox1.Text = "94\n	2.84\n	16.61\n	2.78\n	0.44\n	0.49\n	5.1";
                    break;
                case "도토리묵":
                    richTextBox1.Text = "45\n	0.82\n	8.78\n	0.66\n	0.11\n	0.99\n	0.5";
                    break;
                case "동그랑땡":
                    richTextBox1.Text = "199\n	14.06\n	6.18\n	11.47\n	4.34\n	0.94\n	0.7";
                    break;
                case "동태찌개":
                    richTextBox1.Text = "82\n	1.85\n	4.61\n	12\n	0.31\n	1.05\n	2";
                    break;
                case "된장찌개":
                    richTextBox1.Text = "86\n	2.8\n	8.38\n	7.5\n	0.52\n	1.9\n	1.6";
                    break;
                case "두부김치":
                    richTextBox1.Text = "75\n	3.47\n	6.12\n	6.55\n	0.51\n	1.03\n	0.9";
                    break;
                case "두부조림":
                    richTextBox1.Text = "125\n	9.94\n	2.51\n	6.42\n	1.24\n	0.94\n	0.3";
                    break;
                case "땅콩조림":
                    richTextBox1.Text = "435\n	34.64\n	20.44\n	18.68\n	5.72\n	13.05\n	6.2";
                    break;
                case "떡갈비":
                    richTextBox1.Text = "208\n	14.6\n	2.03\n	16.49\n	5.93\n	0.58\n	0.5";
                    break;
                case "떡국_만두국":
                    richTextBox1.Text = "115\n	1.07\n	22.2\n	3.73\n	0.25\n	0.21\n	0.8";
                    break;
                case "떡꼬치":
                    richTextBox1.Text = "233\n	6.66\n	39.41\n	3.84\n	0.73\n	6.13\n	1.7";
                    break;
                case "떡볶이":
                    richTextBox1.Text = "140\n	1.45\n	30.21\n	3.49\n	0.27\n	3.31\n	2.8";
                    break;
                case "라면":
                    richTextBox1.Text = "453\n	17.1\n	65.5\n	9.3\n	7.63\n	0.73\n	2.4";
                    break;
                case "라볶이":
                    richTextBox1.Text = "139\n	4.61\n	22.57\n	3.44\n	0.95\n	3.32\n	2.4";
                    break;
                case "막국수":
                    richTextBox1.Text = "181\n	2.58\n	35.85\n	6.82\n	0.38\n	1.11\n	0.7";
                    break;
                case "만두":
                    richTextBox1.Text = "112\n	2.64\n	9.56\n	11.55\n	0.74\n	0.57\n	0.5";
                    break;
                case "매운탕":
                    richTextBox1.Text = "71\n	1.77\n	2.08\n	11.37\n	0.25\n	0.6\n	0.7";
                    break;
                case "멍게":
                    richTextBox1.Text = "78\n	2.66\n	3.57\n	9.34\n	0.59\n	0\n	0";
                    break;
                case "메추리알장조림":
                    richTextBox1.Text = "148\n	7.73\n	2.63\n	16.25\n	3.09\n	0.93\n	0.2";
                    break;
                case "멸치볶음":
                    richTextBox1.Text = "173\n	10.22\n	0.87\n	18.58\n	1.96\n	0.19\n	0.3";
                    break;
                case "무국":
                    richTextBox1.Text = "60\n	3.32\n	6.89\n	1.38\n	0.49\n	2.02\n	1.7";
                    break;
                case "무생채":
                    richTextBox1.Text = "40\n	1.14\n	7.17\n	1.15\n	0.18\n	4.03\n	2.2";
                    break;
                case "물냉면":
                    richTextBox1.Text = "96\n	1.44\n	16.14\n	4.13\n	0.49\n	0.47\n	1";
                    break;
                case "물회":
                    richTextBox1.Text = "70\n	1.99\n	6.36\n	6.7\n	0.33\n	0.84\n	0.6";
                    break;
                case "미역국":
                    richTextBox1.Text = "36\n	1.62\n	2.43\n	3.42\n	0.42\n	0.27\n	0.2";
                    break;
                case "미역줄기볶음":
                    richTextBox1.Text = "41\n	1.53\n	6.66\n	0.94\n	0.22\n	2.59\n	0.6";
                    break;
                case "배추김치":
                    richTextBox1.Text = "22\n	0.32\n	4.16\n	1.73\n	0.05\n	1.83\n	1.3";
                    break;
                case "백김치":
                    richTextBox1.Text = "22\n	0.19\n	4.75\n	1.21\n	0.03\n	1.43\n	1.5";
                    break;
                case "보쌈":
                    richTextBox1.Text = "290\n	25.07\n	1.43\n	13.92\n	9.27\n	0.49\n	0.3";
                    break;
                case "부추김치":
                    richTextBox1.Text = "42\n	0.33\n	9.42\n	1.58\n	0.05\n	2.11\n	1.8";
                    break;
                case "북엇국":
                    richTextBox1.Text = "44\n	1.96\n	1.31\n	5.12\n	0.42\n	0.5\n	0.3";
                    break;
                case "불고기":
                    richTextBox1.Text = "162\n	7.95\n	4.92\n	16.13\n	2.02\n	2.39\n	0.9";
                    break;
                case "비빔냉면":
                    richTextBox1.Text = "116\n	2.34\n	18.98\n	4.34\n	0.53\n	0.52\n	0.3";
                    break;
                case "비빔밥":
                    richTextBox1.Text = "146\n	3.5	\n22.44\n	5.51\n	0.93\n	1.5	\n0.7";
                    break;
                case "산낙지":
                    richTextBox1.Text = "60\n	1.32\n	1.57\n	9.77\n	0.24\n	0\n	0";
                    break;
                case "삼겹살":
                    richTextBox1.Text = "331\n	28.21\n	0.59\n	17.24\n	9.67\n	0\n	0";
                    break;
                case "삼계탕":
                    richTextBox1.Text = "91\n	3.18\n	4.08\n	11.1\n	0.64\n	0.09\n	0.2";
                    break;
                case "새우볶음밥":
                    richTextBox1.Text = "162\n	5.86\n	21.15\n	5.53\n	0.99\n	0.76\n	0.7";
                    break;
                case "새우튀김":
                    richTextBox1.Text = "145\n	2.79\n	10.36\n	18.21\n	0.67\n	0.48\n	0.5";
                    break;
                case "생선전":
                    richTextBox1.Text = "150\n	5.74\n	9.57\n	14.12\n	0.96\n	0.38\n	0.5";
                    break;
                case "소세지볶음":
                    richTextBox1.Text = "125\n	9.84\n	2.64\n	6.48\n	2.91\n	0.74\n	0.9";
                    break;
                case "송편":
                    richTextBox1.Text = "231\n	1.49\n	49.07\n	4.57\n	0.23\n	3.2\n	2.2";
                    break;
                case "수육":
                    richTextBox1.Text = "180\n	9.12\n	1.59\n	20.91\n	3.37\n	0.98\n	0.2";
                    break;
                case "수정과":
                    richTextBox1.Text = "48\n	0.15\n	12.89\n	0.28\n	0.01\n	0.02\n	2.9";
                    break;
                case "수제비":
                    richTextBox1.Text = "93\n	1.68\n	15.37\n	3.86\n	0.36\n	0.74\n	1";
                    break;
                case "숙주나물":
                    richTextBox1.Text = "21\n	0.88\n	2.58\n	1.45\n	0.13\n	1.37\n	1";
                    break;
                case "순대":
                    richTextBox1.Text = "158\n	5.01\n	16.18\n	11.16\n	1.75\n	0.07\n	0.4";
                    break;
                case "순두부찌개":
                    richTextBox1.Text = "86	6\n	1.87\n	6.08\n	1.43\n	0.57\n	0.4";
                    break;
                case "시금치나물":
                    richTextBox1.Text = "68\n	3.36\n	7.25\n	3.22\n	0.52\n	0.96\n	2.6";
                    break;
                case "시래기국":
                    richTextBox1.Text = "27\n	1.03\n	3.18\n	2.06\n	0.27\n	0.52\n	0.7";
                    break;
                case "식혜":
                    richTextBox1.Text = "48\n0.02\n12.11\n0.2\n0.01\n9.83\n0";
                    break;
                case "알밥":
                    richTextBox1.Text = "98\n	1.54\n	16.3\n	5.19\n	0.32\n	0.43\n	0.9";
                    break;
                case "애호박볶음":
                    richTextBox1.Text = "119\n	10.24\n	6.99\n	1.6	\n1.46\n	1.69\n	1.9";
                    break;
                case "약과":
                    richTextBox1.Text = "407\n	13.31\n	66.24\n	7\n	1.33\n	22.27\n	1.9";
                    break;
                case "약식":
                    richTextBox1.Text = "255\n	2.79\n	53.83\n	3.51\n	0.31\n	15.12\n	1.7";
                    break;
                case "양념게장":
                    richTextBox1.Text = "134\n	5.57\n	7.32\n	15.03\n	0.82\n	1.81\n	2.8";
                    break;
                case "양념치킨":
                    richTextBox1.Text = "288\n	18.01\n	13.72\n	16.44\n	5.2\n	0.1\n	0.6";
                    break;
                case "어묵볶음":
                    richTextBox1.Text = "125\n	4.9\n	14.2\n	7.19\n	0.6\n	1.47\n	1.9";
                    break;
                case "연근조림":
                    richTextBox1.Text = "96\n	1.21\n	20.13\n	2.84\n	0.18\n	4.84\n	3.7";
                    break;
                case "열무국수":
                    richTextBox1.Text = "112\n	2.31\n	19.52\n	3\n	0.33\n	0.63\n	0.5";
                    break;
                case "열무김치":
                    richTextBox1.Text = "32\n	0.28\n	6.86\n	1.48\n	0.05\n	2.79\n	2.1";
                    break;
                case "오이소박이":
                    richTextBox1.Text = "27\n	0.38\n	5.39\n	1.41\n	0.08\n	2.22\n	1.2";
                    break;
                case "오징어채볶음":
                    richTextBox1.Text = "122\n	4.36\n	8.56\n	12.09\n	0.99\n	1.87\n	1.2";
                    break;
                case "오징어튀김":
                    richTextBox1.Text = "174\n	5.29\n	16.09	\n14.51\n	0.82\n	1.22\n	0.9";
                    break;
                case "우엉조림":
                    richTextBox1.Text = "90\n	1.95\n	16.54\n	1.96\n	0.28\n	5.12\n	1.4";
                    break;
                case "유부초밥":
                    richTextBox1.Text = "180\n	6.42\n	24.37\n	7.09\n	0.95\n	1.13\n	1.7";
                    break;
                case "육개장":
                    richTextBox1.Text = "83\n	3.78\n	5.23\n	7.16\n	1.04\n	0.73\n	1.3";
                    break;
                case "육회":
                    richTextBox1.Text = "112\n	4.87\n	2.37\n	13.95\n	1.81\n	0.97\n	0.4";
                    break;
                case "잔치국수":
                    richTextBox1.Text = "89\n	0.79\n	16.63\n	4.16\n	0.2	\n0.53\n	1.1";
                    break;
                case "잡곡밥":
                    richTextBox1.Text = "162\n	0.94\n	34.99\n	4.2\n	0.16\n	0.05\n	1.8";
                    break;
                case "잡채":
                    richTextBox1.Text = "190\n	6.78\n	28.82\n	5.06\n	0.96\n	2.02\n	2.3";
                    break;
                case "장어구이":
                    richTextBox1.Text = "224\n	13.09\n	5.95\n	19.59\n	2.6\n	1.21\n	0.9";
                    break;
                case "장조림":
                    richTextBox1.Text = "147\n	7.38\n	2.86\n	16.68\n	3.06\n	0.97\n	0.2";
                    break;
                case "전복죽":
                    richTextBox1.Text = "57\n	0.7\n	11.05\n	1.47\n	0.74\n	0\n	0";
                    break;
                case "젓갈":
                    richTextBox1.Text = "128\n	3\n	12.81\n	11.74\n	0.49\n	0.44\n	0.8";
                    break;
                case "제육볶음":
                    richTextBox1.Text = "191\n	11.03\n	10.37\n	13.21\n	3.33\n	6.98\n	1.6";
                    break;
                case "조개구이":
                    richTextBox1.Text = "138\n	6.86\n	3.15\n	14.95\n	1.18\n	0\n	0";
                    break;
                case "조기구이":
                    richTextBox1.Text = "132\n	4.01\n	0	\n22.51\n	1.38\n	0\n	0";
                    break;
                case "족발":
                    richTextBox1.Text = "229\n	15.71\n	0.32\n	21.53\n	4.25\n	0.02\n	0";
                    break;
                case "주꾸미볶음":
                    richTextBox1.Text = "188\n	7.12\n	10.52\n	20.53\n	1.14\n	2.51\n	1.9";
                    break;
                case "주먹밥":
                    richTextBox1.Text = "188\n	3.09\n	34.49\n	4.84\n	0.47\n	0.38\n	1.5";
                    break;
                case "짜장면":
                    richTextBox1.Text = "175\n	4.44\n	28.76\n	5.96\n	1.25\n	1.46\n	3.1";
                    break;
                case "짬뽕":
                    richTextBox1.Text = "153\n	2.4\n	26.71\n	5.76\n	0.3\n	0.65\n	0.6";
                    break;
                case "쫄면":
                    richTextBox1.Text = "106\n	0.77\n	21.64\n	2.27\n	0.12\n	4.17\n	2.2";
                    break;
                case "찜닭":
                    richTextBox1.Text = "193\n	7.69\n	15.82\n	13.93\n	1.74\n	3.76\n	1.2";
                    break;
                case "총각김치":
                    richTextBox1.Text = "29\n	0.26\n	6.18\n	1.22\n	0.05\n	2.83\n	1.9";
                    break;
                case "추어탕":
                    richTextBox1.Text = "65\n	2.67\n	1.18\n	8.84\n	0.52\n	0.33\n	0.5";
                    break;
                case "칼국수":
                    richTextBox1.Text = "120\n	2.51\n	20.5\n	4.58\n	0.45\n	1.02\n	2.1";
                    break;
                case "코다리조림":
                    richTextBox1.Text = "110\n2.57\n5.74\n14.54\n0.36\n3.3\n0.7";
                    break;
                case "콩국수":
                    richTextBox1.Text = "256\n5.24\n40.7\n12.5\n1.03\n1.1\n3.1";
                    break;
                case "콩나물국":
                    richTextBox1.Text = "16\n0.62\n2.35\n0.87\n0.1\n0.7\n0.8";
                    break;
                case "콩나물무침":
                    richTextBox1.Text = "48\n1.41\n7.49\n2.89\n0.23\n2.46\n2.4";
                    break;
                case "콩자반":
                    richTextBox1.Text = "317\n1.37\n60.56\n17.78\n0.32\n9.18\n12";
                    break;
                case "파김치":
                    richTextBox1.Text = "38\n0.31\n8.24\n2.01\n0.05\n2.72\n2.5";
                    break;
                case "파전":
                    richTextBox1.Text = "136\n5.61\n16.79\n5.15\n1.15\n3.54\n1.8";
                    break;
                case "편육":
                    richTextBox1.Text = "163\n8.5\n3.06\n17.43\n2.82\n0\n0.6";
                    break;
                case "피자":
                    richTextBox1.Text = "244\n10.58\n26.34\n10.81\n3.99\n3.94\n2.4";
                    break;
                case "해물찜":
                    richTextBox1.Text = "135\n5.76\n8.92\n11.85\n0.92\n1.09\n2.2";
                    break;
                case "호박전":
                    richTextBox1.Text = "95\n6.79\n6.7\n2.37\n0.82\n1.12\n0.9";
                    break;
                case "호박죽":
                    richTextBox1.Text = "70\n0.99\n14.67\n1.46\n0.58\n1.82\n1.1";
                    break;
                case "홍어무침":
                    richTextBox1.Text = "87\n2.11\n5.71\n12.31\n0.4\n1.55\n2";
                    break;
                case "황태구이":
                    richTextBox1.Text = "294\n13.48\n6.08\n36.68\n2.69\n2.33\n2.2";
                    break;
                case "회무침":
                    richTextBox1.Text = "95\n3.14\n14.42\n5.61\n0.55\n3.4\n4.1";
                    break;
                case "후라이드치킨":
                    richTextBox1.Text = "290\n17.27\n9.98\n22.31\n4.32\n0.23\n0.3";
                    break;
                case "훈제오리":
                    richTextBox1.Text = "308\n24.8\n0\n19.79\n8.53\n0\n0";
                    break;
            }
        }
    }
}
