using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;

// Добавленные using.
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Diagnostics;

namespace NF
{
    public partial class MainWindow : Window
    {
        // Загрузка окна.
        public MainWindow()
        {
            InitializeComponent();

            AddMarker.IsEnabled = false;
            Map.IsEnabled = false;
            Map.Opacity = 0.5;
            ScrollViewer.IsEnabled = false;
        }

        // Вычисление координат.
        private void Map_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Point point = e.GetPosition(Map);

            double xDiff = 1633 - 750;
            double yDiff = 833 - 80;

            int newX = (int)(750 + (point.X * xDiff / Map.ActualWidth));
            int newY = (int)(80 + (point.Y * yDiff / Map.ActualHeight));

            X.Text = newX.ToString();
            Y.Text = newY.ToString();
        }

        // Масштабирование карты.
        private void Map_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            double zoomFactor = 1.1;

            if (e.Delta > 0)
            {
                Map.Width *= zoomFactor;
                Map.Height *= zoomFactor;
            }
            else
            {
                Map.Width /= zoomFactor;
                Map.Height /= zoomFactor;
            }

            e.Handled = true;
        }

        // Счетчики загрузки.
        private static int currentRowIndexAdm = 3;
        private static int currentRowIndexMed = 3;
        private static int currentRowIndexProd = 3;
        private static int currentRowIndexEduc = 3;
        private static int currentRowIndexRest = 3;
        private static int currentRowIndexChill = 3;
        private static int currentRowIndexTransp = 3;
        private static int currentRowIndexEnter = 3;
        private static int currentRowIndexHouse = 3;
        private static int currentRowIndexGoods = 3;

        // Уведомления.
        private bool administration = true;
        private bool medecine = true;
        private bool products = true;
        private bool education = true;
        private bool restaurants = true;
        private bool chillzone = true;
        private bool transport = true;
        private bool entertainment = true;
        private bool houseservices = true;
        private bool goods = true;

        // Строки подключения.
        private static string folderPath = "";
        private static string htmlFilePath;
        private static string cssFilePath;
        private static string jsFilePath;

        // Выбор папки с сайтом.
        private void ChooseSite_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new OpenFileDialog
            {
                Title = "Выберите папку сайта",
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "Путь",
                Filter = "Folders|no.files",
                ValidateNames = false
            };

            if (folderDialog.ShowDialog() == true)
            {
                folderPath = Path.GetDirectoryName(folderDialog.FileName);
                htmlFilePath = Path.Combine(folderPath, "index.html");
                cssFilePath = Path.Combine(folderPath, "styles.css");
                jsFilePath = Path.Combine(folderPath, "script.js");
            }
        }

        // Загрузка excel-файла и перелистывание.
        private void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            if (folderPath == "")
                MessageBox.Show("Выберите папку сайта!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            else
            {
                Process[] processes = Process.GetProcessesByName("EXCEL");
                if (processes.Length > 0)
                    MessageBox.Show("Пожалуйста, закройте Excel-файл!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                else
                {
                    ScrollViewer.IsEnabled = true;
                    Map.IsEnabled = true;
                    Map.Opacity = 1;
                    LoadExcel.Content = "Продолжить";
                    LoadExcel.IsEnabled = false;

                    string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    string excelFilePath = Path.Combine(appDirectory, "Данные.xlsx");

                    if (File.Exists(excelFilePath))
                    {
                        FileInfo fileInfo = new FileInfo(excelFilePath);
                        using (ExcelPackage package = new ExcelPackage(fileInfo))
                        {
                            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                            if (worksheet != null)
                            {
                                if (administration)
                                    MessageBox.Show("Администрация и госуслги", "Администрация и госуслуги", MessageBoxButton.OK, MessageBoxImage.Information);
                                administration = false;

                                for (int rowIndex1 = currentRowIndexAdm; rowIndex1 <= 100; rowIndex1++)
                                {
                                    object F = worksheet.Cells[$"F{rowIndex1}"].Value;

                                    if (F == null)
                                    {
                                        object A1 = worksheet.Cells[$"A{rowIndex1}"].Value;
                                        object A2 = worksheet.Cells[$"B{rowIndex1}"].Value;
                                        object A3 = worksheet.Cells[$"C{rowIndex1}"].Value;
                                        object A4 = worksheet.Cells[$"D{rowIndex1}"].Value;
                                        object A5 = worksheet.Cells[$"E{rowIndex1}"].Value;

                                        if (A1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = A1?.ToString();
                                            Name.Text = A2?.ToString();
                                            Special.Text = A3?.ToString();
                                            Adress.Text = A4?.ToString();
                                            Site.Text = A5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexAdm = rowIndex1 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                if (medecine)
                                    MessageBox.Show("Медицина", "Медицина", MessageBoxButton.OK, MessageBoxImage.Information);
                                medecine = false;

                                for (int rowIndex2 = currentRowIndexMed; rowIndex2 <= 100; rowIndex2++)
                                {
                                    object L = worksheet.Cells[$"L{rowIndex2}"].Value;

                                    if (L == null)
                                    {
                                        object M1 = worksheet.Cells[$"G{rowIndex2}"].Value;
                                        object M2 = worksheet.Cells[$"H{rowIndex2}"].Value;
                                        object M3 = worksheet.Cells[$"I{rowIndex2}"].Value;
                                        object M4 = worksheet.Cells[$"J{rowIndex2}"].Value;
                                        object M5 = worksheet.Cells[$"K{rowIndex2}"].Value;

                                        if (M1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = M1?.ToString();
                                            Name.Text = M2?.ToString();
                                            Special.Text = M3?.ToString();
                                            Adress.Text = M4?.ToString();
                                            Site.Text = M5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexMed = rowIndex2 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                if (products)
                                    MessageBox.Show("Продукты", "Продукты", MessageBoxButton.OK, MessageBoxImage.Information);
                                products = false;

                                for (int rowIndex3 = currentRowIndexProd; rowIndex3 <= 100; rowIndex3++)
                                {
                                    object R = worksheet.Cells[$"R{rowIndex3}"].Value;

                                    if (R == null)
                                    {
                                        object P1 = worksheet.Cells[$"M{rowIndex3}"].Value;
                                        object P2 = worksheet.Cells[$"N{rowIndex3}"].Value;
                                        object P3 = worksheet.Cells[$"O{rowIndex3}"].Value;
                                        object P4 = worksheet.Cells[$"P{rowIndex3}"].Value;
                                        object P5 = worksheet.Cells[$"Q{rowIndex3}"].Value;

                                        if (P1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = P1?.ToString();
                                            Name.Text = P2?.ToString();
                                            Special.Text = P3?.ToString();
                                            Adress.Text = P4?.ToString();
                                            Site.Text = P5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexProd = rowIndex3 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                if (education)
                                    MessageBox.Show("Образование", "Образование", MessageBoxButton.OK, MessageBoxImage.Information);
                                education = false;

                                for (int rowIndex4 = currentRowIndexEduc; rowIndex4 <= 100; rowIndex4++)
                                {
                                    object X_ = worksheet.Cells[$"X{rowIndex4}"].Value;

                                    if (X_ == null)
                                    {
                                        object E1 = worksheet.Cells[$"S{rowIndex4}"].Value;
                                        object E2 = worksheet.Cells[$"T{rowIndex4}"].Value;
                                        object E3 = worksheet.Cells[$"U{rowIndex4}"].Value;
                                        object E4 = worksheet.Cells[$"V{rowIndex4}"].Value;
                                        object E5 = worksheet.Cells[$"W{rowIndex4}"].Value;

                                        if (E1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = E1?.ToString();
                                            Name.Text = E2?.ToString();
                                            Special.Text = E3?.ToString();
                                            Adress.Text = E4?.ToString();
                                            Site.Text = E5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexEduc = rowIndex4 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                if (restaurants)
                                    MessageBox.Show("Кафе и рестораны", "Кафе и рестораны", MessageBoxButton.OK, MessageBoxImage.Information);
                                restaurants = false;

                                for (int rowIndex5 = currentRowIndexRest; rowIndex5 <= 100; rowIndex5++)
                                {
                                    object AD = worksheet.Cells[$"AD{rowIndex5}"].Value;

                                    if (AD == null)
                                    {
                                        object R1 = worksheet.Cells[$"Y{rowIndex5}"].Value;
                                        object R2 = worksheet.Cells[$"Z{rowIndex5}"].Value;
                                        object R3 = worksheet.Cells[$"AA{rowIndex5}"].Value;
                                        object R4 = worksheet.Cells[$"AB{rowIndex5}"].Value;
                                        object R5 = worksheet.Cells[$"AC{rowIndex5}"].Value;

                                        if (R1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = R1?.ToString();
                                            Name.Text = R2?.ToString();
                                            Special.Text = R3?.ToString();
                                            Adress.Text = R4?.ToString();
                                            Site.Text = R5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexRest = rowIndex5 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                if (chillzone)
                                    MessageBox.Show("Зоны отдыха", "Зоны отдыха", MessageBoxButton.OK, MessageBoxImage.Information);
                                chillzone = false;

                                for (int rowIndex6 = currentRowIndexChill; rowIndex6 <= 100; rowIndex6++)
                                {
                                    object AJ = worksheet.Cells[$"AJ{rowIndex6}"].Value;

                                    if (AJ == null)
                                    {
                                        object CZ1 = worksheet.Cells[$"AE{rowIndex6}"].Value;
                                        object CZ2 = worksheet.Cells[$"AF{rowIndex6}"].Value;
                                        object CZ3 = worksheet.Cells[$"AG{rowIndex6}"].Value;
                                        object CZ4 = worksheet.Cells[$"AH{rowIndex6}"].Value;
                                        object CZ5 = worksheet.Cells[$"AI{rowIndex6}"].Value;

                                        if (CZ1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = CZ1?.ToString();
                                            Name.Text = CZ2?.ToString();
                                            Special.Text = CZ3?.ToString();
                                            Adress.Text = CZ4?.ToString();
                                            Site.Text = CZ5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexChill = rowIndex6 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                if (transport)
                                    MessageBox.Show("Транспорт", "Транспорт", MessageBoxButton.OK, MessageBoxImage.Information);
                                transport = false;

                                for (int rowIndex7 = currentRowIndexTransp; rowIndex7 <= 100; rowIndex7++)
                                {
                                    object AP = worksheet.Cells[$"AP{rowIndex7}"].Value;

                                    if (AP == null)
                                    {
                                        object T1 = worksheet.Cells[$"AK{rowIndex7}"].Value;
                                        object T2 = worksheet.Cells[$"AL{rowIndex7}"].Value;
                                        object T3 = worksheet.Cells[$"AM{rowIndex7}"].Value;
                                        object T4 = worksheet.Cells[$"AN{rowIndex7}"].Value;
                                        object T5 = worksheet.Cells[$"AO{rowIndex7}"].Value;

                                        if (T1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = T1?.ToString();
                                            Name.Text = T2?.ToString();
                                            Special.Text = T3?.ToString();
                                            Adress.Text = T4?.ToString();
                                            Site.Text = T5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexTransp = rowIndex7 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                if (entertainment)
                                    MessageBox.Show("Учреждения культуры и досуга", "Учреждения культуры и досуга", MessageBoxButton.OK, MessageBoxImage.Information);
                                entertainment = false;

                                for (int rowIndex8 = currentRowIndexEnter; rowIndex8 <= 100; rowIndex8++)
                                {
                                    object AV = worksheet.Cells[$"AV{rowIndex8}"].Value;

                                    if (AV == null)
                                    {
                                        object En1 = worksheet.Cells[$"AQ{rowIndex8}"].Value;
                                        object En2 = worksheet.Cells[$"AR{rowIndex8}"].Value;
                                        object En3 = worksheet.Cells[$"AS{rowIndex8}"].Value;
                                        object En4 = worksheet.Cells[$"AT{rowIndex8}"].Value;
                                        object En5 = worksheet.Cells[$"AU{rowIndex8}"].Value;

                                        if (En1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = En1?.ToString();
                                            Name.Text = En2?.ToString();
                                            Special.Text = En3?.ToString();
                                            Adress.Text = En4?.ToString();
                                            Site.Text = En5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexEnter = rowIndex8 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                if (houseservices)
                                    MessageBox.Show("Бытовые услуги", "Бытовые услуги", MessageBoxButton.OK, MessageBoxImage.Information);
                                houseservices = false;

                                for (int rowIndex9 = currentRowIndexHouse; rowIndex9 <= 100; rowIndex9++)
                                {
                                    object BB = worksheet.Cells[$"BB{rowIndex9}"].Value;

                                    if (BB == null)
                                    {
                                        object H1 = worksheet.Cells[$"AW{rowIndex9}"].Value;
                                        object H2 = worksheet.Cells[$"AX{rowIndex9}"].Value;
                                        object H3 = worksheet.Cells[$"AY{rowIndex9}"].Value;
                                        object H4 = worksheet.Cells[$"AZ{rowIndex9}"].Value;
                                        object H5 = worksheet.Cells[$"BA{rowIndex9}"].Value;

                                        if (H1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = H1?.ToString();
                                            Name.Text = H2?.ToString();
                                            Special.Text = H3?.ToString();
                                            Adress.Text = H4?.ToString();
                                            Site.Text = H5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexHouse = rowIndex9 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                if (goods)
                                    MessageBox.Show("Промтовары", "Промтовары", MessageBoxButton.OK, MessageBoxImage.Information);
                                goods = false;

                                for (int rowIndex10 = currentRowIndexGoods; rowIndex10 <= 100; rowIndex10++)
                                {
                                    object BH = worksheet.Cells[$"BH{rowIndex10}"].Value;

                                    if (BH == null)
                                    {
                                        object G1 = worksheet.Cells[$"BC{rowIndex10}"].Value;
                                        object G2 = worksheet.Cells[$"BD{rowIndex10}"].Value;
                                        object G3 = worksheet.Cells[$"BE{rowIndex10}"].Value;
                                        object G4 = worksheet.Cells[$"BF{rowIndex10}"].Value;
                                        object G5 = worksheet.Cells[$"BG{rowIndex10}"].Value;

                                        if (G1 != null)
                                        {
                                            // Вставляем значения.
                                            PhotoNum.Text = G1?.ToString();
                                            Name.Text = G2?.ToString();
                                            Special.Text = G3?.ToString();
                                            Adress.Text = G4?.ToString();
                                            Site.Text = G5?.ToString();

                                            X.Text = "";
                                            Y.Text = "";
                                            AddMarker.IsEnabled = false;

                                            currentRowIndexGoods = rowIndex10 + 1;

                                            return;
                                        }
                                        else
                                            break;
                                    }
                                }

                                LoadExcel.Content = "Загрузить Excel";
                                MessageBox.Show("Маркеры закончились!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);

                                PhotoNum.Text = "";
                                Name.Text = "";
                                Special.Text = "";
                                Adress.Text = "";
                                Site.Text = "";
                                X.Text = "";
                                Y.Text = "";

                                currentRowIndexAdm = 3;
                                currentRowIndexMed = 3;
                                currentRowIndexProd = 3;
                                currentRowIndexEduc = 3;
                                currentRowIndexRest = 3;
                                currentRowIndexChill = 3;
                                currentRowIndexTransp = 3;
                                currentRowIndexEnter = 3;
                                currentRowIndexHouse = 3;
                                currentRowIndexGoods = 3;

                                // administration = true;
                                medecine = true;
                                products = true;
                                education = true;
                                restaurants = true;
                                chillzone = true;
                                transport = true;
                                entertainment = true;
                                houseservices = true;
                                goods = true;

                                LoadExcel.IsEnabled = false;
                                AddMarker.IsEnabled = false;
                                Map.IsEnabled = false;
                                Map.Opacity = 0.5;
                            }
                        }
                    }
                }
            }           
        }

        // Счетчики добавления.
        private static int currentRowIndexAdm_ = 3;
        private static int currentRowIndexMed_ = 3;
        private static int currentRowIndexProd_ = 3;
        private static int currentRowIndexEduc_ = 3;
        private static int currentRowIndexRest_ = 3;
        private static int currentRowIndexChill_ = 3;
        private static int currentRowIndexTransp_ = 3;
        private static int currentRowIndexEnter_ = 3;
        private static int currentRowIndexHouse_ = 3;
        private static int currentRowIndexGoods_ = 3;

        // Добавление маркера.
        private void AddMarker_Click(object sender, RoutedEventArgs e)
        {
            LoadExcel.IsEnabled = true;

            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string excelFilePath = Path.Combine(appDirectory, "Данные.xlsx");

            FileInfo fileInfo = new FileInfo(excelFilePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet != null)
                {
                    // Администрация и госуслуги.
                    for (int rowIndex1 = currentRowIndexAdm_; rowIndex1 <= 100; rowIndex1++)
                    {
                        object F = worksheet.Cells[$"F{rowIndex1}"].Value;
                        object A1 = worksheet.Cells[$"A{rowIndex1}"].Value;

                        if (F == null && A1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""adm-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""admin-sidebar{PhotoNum_}"">
        <img src=""Administration/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#adm-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #admin-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#admin-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#admin-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#admin-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#admin-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#admin-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('adm-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker a1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('admin-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker a2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"F{rowIndex1}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexAdm_ = rowIndex1 + 1;

                            return;
                        }
                    }

                    // Медицина.
                    for (int rowIndex2 = currentRowIndexMed_; rowIndex2 <= 100; rowIndex2++)
                    {
                        object L = worksheet.Cells[$"L{rowIndex2}"].Value;
                        object M1 = worksheet.Cells[$"G{rowIndex2}"].Value;

                        if (L == null && M1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""med-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""medicine-sidebar{PhotoNum_}"" class=""sidebar"">
        <img src=""Medicine/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#med-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #medicine-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#medicine-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#medicine-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#medicine-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#medicine-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#medicine-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('med-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker m1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('medicine-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker m2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"L{rowIndex2}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexMed_ = rowIndex2 + 1;

                            return;
                        }
                    }

                    // Продукты.
                    for (int rowIndex3 = currentRowIndexProd_; rowIndex3 <= 100; rowIndex3++)
                    {
                        object R = worksheet.Cells[$"R{rowIndex3}"].Value;
                        object P1 = worksheet.Cells[$"M{rowIndex3}"].Value;

                        if (R == null && P1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""mar-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""market-sidebar{PhotoNum_}"" class=""sidebar"">
        <img src=""Products/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#mar-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #market-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#market-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#market-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#market-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#market-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#market-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('mar-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker p1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('market-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker p2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"R{rowIndex3}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexProd_ = rowIndex3 + 1;

                            return;
                        }
                    }

                    // Образование.
                    for (int rowIndex4 = currentRowIndexEduc_; rowIndex4 <= 100; rowIndex4++)
                    {
                        object X_ = worksheet.Cells[$"X{rowIndex4}"].Value;
                        object E1 = worksheet.Cells[$"S{rowIndex4}"].Value;

                        if (X_ == null && E1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""edu-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""education-sidebar{PhotoNum_}"" class=""sidebar"">
        <img src=""Education/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#edu-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #education-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#education-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#education-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#education-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#education-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#education-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('edu-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker ed1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('education-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker ed2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"X{rowIndex4}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexEduc_ = rowIndex4 + 1;

                            return;
                        }
                    }

                    // Рестораны.
                    for (int rowIndex5 = currentRowIndexRest_; rowIndex5 <= 100; rowIndex5++)
                    {
                        object AD = worksheet.Cells[$"AD{rowIndex5}"].Value;
                        object R1 = worksheet.Cells[$"Y{rowIndex5}"].Value;

                        if (AD == null && R1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""res-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""restaurant-sidebar{PhotoNum_}"" class=""sidebar"">
        <img src=""Restaurants/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#res-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #restaurant-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#restaurant-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#restaurant-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#restaurant-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#restaurant-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#restaurant-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('res-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker r1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('restaurant-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker r2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"AD{rowIndex5}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexRest_ = rowIndex5 + 1;

                            return;
                        }
                    }

                    // Зоны отдыха.
                    for (int rowIndex6 = currentRowIndexChill_; rowIndex6 <= 100; rowIndex6++)
                    {
                        object AJ = worksheet.Cells[$"AJ{rowIndex6}"].Value;
                        object CZ1 = worksheet.Cells[$"AE{rowIndex6}"].Value;

                        if (AJ == null && CZ1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""par-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""park-sidebar{PhotoNum_}"" class=""sidebar"">
        <img src=""Park/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#par-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #park-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#park-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#park-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#park-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#park-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#park-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('par-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker c1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('park-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker c2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"AJ{rowIndex6}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexChill_ = rowIndex6 + 1;

                            return;
                        }
                    }

                    // Транспорт.
                    for (int rowIndex7 = currentRowIndexTransp_; rowIndex7 <= 100; rowIndex7++)
                    {
                        object AP = worksheet.Cells[$"AP{rowIndex7}"].Value;
                        object T1 = worksheet.Cells[$"AK{rowIndex7}"].Value;

                        if (AP == null && T1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""tra-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""transport-sidebar{PhotoNum_}"" class=""sidebar"">
        <img src=""Transport/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#tra-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #transport-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#transport-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#transport-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#transport-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#transport-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#transport-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('tra-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker t1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('transport-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker t2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"AP{rowIndex7}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexTransp_ = rowIndex7 + 1;

                            return;
                        }
                    }

                    // Учреждения культуры и досуга.
                    for (int rowIndex8 = currentRowIndexEnter_; rowIndex8 <= 100; rowIndex8++)
                    {
                        object AV = worksheet.Cells[$"AV{rowIndex8}"].Value;
                        object En1 = worksheet.Cells[$"AQ{rowIndex8}"].Value;

                        if (AV == null && En1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""ent-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""entertainment-sidebar{PhotoNum_}"" class=""sidebar"">
        <img src=""Entertaintments/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#ent-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #entertainment-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#entertainment-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#entertainment-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#entertainment-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#entertainment-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#entertainment-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('ent-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker e1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('entertainment-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker e2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"AV{rowIndex8}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexEnter_ = rowIndex8 + 1;

                            return;
                        }
                    }

                    // Бытовые услуги.
                    for (int rowIndex9 = currentRowIndexHouse_; rowIndex9 <= 100; rowIndex9++)
                    {
                        object BB = worksheet.Cells[$"BB{rowIndex9}"].Value;
                        object H1 = worksheet.Cells[$"AW{rowIndex9}"].Value;

                        if (BB == null && H1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""ser-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""service-sidebar{PhotoNum_}"" class=""sidebar"">
        <img src=""Services/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#ser-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #service-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#service-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#service-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#service-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#service-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#service-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('ser-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker h1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('service-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker h2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"BB{rowIndex9}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexHouse_ = rowIndex9 + 1;

                            return;
                        }
                    }

                    // Промтовары.
                    for (int rowIndex10 = currentRowIndexGoods_; rowIndex10 <= 100; rowIndex10++)
                    {
                        object BH = worksheet.Cells[$"BH{rowIndex10}"].Value;
                        object G1 = worksheet.Cells[$"BC{rowIndex10}"].Value;

                        if (BH == null && G1 != null)
                        {
                            string PhotoNum_ = PhotoNum.Text;
                            string Name_ = Name.Text;
                            string Special_ = Special.Text;
                            string Adress_ = Adress.Text;
                            string Site_ = Site.Text;
                            string _X = X.Text;
                            string _Y = Y.Text;

                            // Добавление маркера в HTML.
                            List<string> htmlLines1 = File.ReadAllLines(htmlFilePath).ToList();
                            string htmlToInsert1 = $@"        <button id=""pro-marker-button{PhotoNum_}"" style=""display: none;""></button>";
                            string searchHtmlMarker1 = "<!-- MARKER 1 -->";
                            int htmlMarkerIndex1 = htmlLines1.FindIndex(line => line.Contains(searchHtmlMarker1));
                            if (htmlMarkerIndex1 != -1)
                                htmlLines1.Insert(htmlMarkerIndex1 + 1, htmlToInsert1);
                            File.WriteAllLines(htmlFilePath, htmlLines1);

                            List<string> htmlLines2 = File.ReadAllLines(htmlFilePath).ToList();

                            string htmlToInsert2 = $@"
    <div id=""prom-sidebar{PhotoNum_}"" class=""sidebar"">
        <img src=""Prom/{PhotoNum_}.jpg"" alt=""Brateevo"">
        <div id=""brateevo-text"" class=""textbox bold"">{Name_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">ДОСТУПНОСТЬ</div>
        <div class=""textbox small-gap"">{Special_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">АДРЕС</div>
        <div class=""textbox small-gap"">{Adress_}</div>
        <div class=""textbox bold"" style=""color: #c90000;"">САЙТ</div>
        <div class=""textbox small-gap""><a href=""https://{Site_}"">{Site_}</a></div>
    </div>";

                            string searchHtmlMarker2 = "<!-- MARKER 2 -->";
                            int htmlMarkerIndex2 = htmlLines2.FindIndex(line => line.Contains(searchHtmlMarker2));
                            if (htmlMarkerIndex2 != -1)
                                htmlLines2.Insert(htmlMarkerIndex2 + 1, htmlToInsert2);
                            File.WriteAllLines(htmlFilePath, htmlLines2);

                            // Добавление стилей в CSS.
                            List<string> cssLines1 = File.ReadAllLines(cssFilePath).ToList();
                            int cssLineNumber1 = 23;

                            cssLines1.Insert(cssLineNumber1, $@"
#pro-marker-button{PhotoNum_} {{
    position: absolute;
    top: {_Y}px;
    left: {_X}px;
    width: 40px;
    height: 40px;
    background-image: url(marker.png);
    background-repeat: no-repeat;
    background-size: contain;
    background-color: transparent;
    border: none;
    outline: none;
    cursor: pointer;
}}");

                            File.WriteAllLines(cssFilePath, cssLines1);

                            string insertMarker1 = "/* === INSERT_MARKER_1 === */";
                            string cssToInsert1 = $@"  #prom-sidebar{PhotoNum_},";
                            List<string> cssLines2 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex1 = cssLines2.IndexOf(insertMarker1);
                            if (cssMarkerIndex1 != -1)
                                cssLines2.Insert(cssMarkerIndex1 + 1, cssToInsert1);
                            File.WriteAllLines(cssFilePath, cssLines2);

                            string insertMarker2 = "/* === INSERT_MARKER_2 === */";
                            string cssToInsert2 = $@"#prom-sidebar{PhotoNum_}.active,";
                            List<string> cssLines3 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex2 = cssLines3.IndexOf(insertMarker2);
                            if (cssMarkerIndex2 != -1)
                                cssLines3.Insert(cssMarkerIndex2 + 1, cssToInsert2);
                            File.WriteAllLines(cssFilePath, cssLines3);

                            string insertMarker3 = "/* === INSERT_MARKER_3 === */";
                            string cssToInsert3 = $@"#prom-sidebar{PhotoNum_} img,";
                            List<string> cssLines4 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex3 = cssLines4.IndexOf(insertMarker3);
                            if (cssMarkerIndex3 != -1)
                                cssLines4.Insert(cssMarkerIndex3 + 1, cssToInsert3);
                            File.WriteAllLines(cssFilePath, cssLines4);

                            string insertMarker4 = "/* === INSERT_MARKER_4 === */";
                            string cssToInsert4 = $@"#prom-sidebar{PhotoNum_} .textbox,";
                            List<string> cssLines5 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex4 = cssLines5.IndexOf(insertMarker4);
                            if (cssMarkerIndex4 != -1)
                                cssLines5.Insert(cssMarkerIndex4 + 1, cssToInsert4);
                            File.WriteAllLines(cssFilePath, cssLines5);

                            string insertMarker5 = "/* === INSERT_MARKER_5 === */";
                            string cssToInsert5 = $@"#prom-sidebar{PhotoNum_} .bold,";
                            List<string> cssLines6 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex5 = cssLines6.IndexOf(insertMarker5);
                            if (cssMarkerIndex5 != -1)
                                cssLines6.Insert(cssMarkerIndex5 + 1, cssToInsert5);
                            File.WriteAllLines(cssFilePath, cssLines6);

                            string insertMarker6 = "/* === INSERT_MARKER_6 === */";
                            string cssToInsert6 = $@"#prom-sidebar{PhotoNum_} .small-gap,";
                            List<string> cssLines7 = File.ReadAllLines(cssFilePath).ToList();
                            int cssMarkerIndex6 = cssLines7.IndexOf(insertMarker6);
                            if (cssMarkerIndex6 != -1)
                                cssLines7.Insert(cssMarkerIndex6 + 1, cssToInsert6);
                            File.WriteAllLines(cssFilePath, cssLines7);

                            // Добавление скриптов в JavaScript.
                            List<string> jsLines1 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert1 = $@"        document.getElementById('pro-marker-button{PhotoNum_}'),";
                            string searchMarkerA1 = "// Marker g1";
                            int markerIndexA1 = jsLines1.FindIndex(line => line.Contains(searchMarkerA1));
                            if (markerIndexA1 != -1)
                                jsLines1.Insert(markerIndexA1 + 1, codeToInsert1);
                            File.WriteAllLines(jsFilePath, jsLines1);

                            List<string> jsLines2 = File.ReadAllLines(jsFilePath).ToList();
                            string codeToInsert2 = $@"        document.getElementById('prom-sidebar{PhotoNum_}'),";
                            string searchMarkerA2 = "// Marker g2";
                            int markerIndexA2 = jsLines2.FindIndex(line => line.Contains(searchMarkerA2));
                            if (markerIndexA2 != -1)
                                jsLines2.Insert(markerIndexA2 + 1, codeToInsert2);
                            File.WriteAllLines(jsFilePath, jsLines2);

                            // Добавление + в Excel.
                            worksheet.Cells[$"BH{rowIndex10}"].Value = "+";
                            package.Save();

                            AddMarker.IsEnabled = false;
                            ScrollViewer.IsEnabled = false;
                            Map.IsEnabled = false;
                            Map.Opacity = 0.5;

                            currentRowIndexGoods_ = rowIndex10 + 1;

                            return;
                        }
                    }
                }
            }
        }

        // Активация добавления маркера при изменении коордиинат.
        private void X_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            AddMarker.IsEnabled = true;
        }

        // Срабатывание кнопки Добавить по нажатию enter.
        private void AddMarker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddMarker_Click(sender, e);
            }
        }

        // Завершение работы программы.
        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // Возможность перетаскивания программы.
        private Point startPoint;
        private void TitleBar_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            startPoint = e.GetPosition(this);
        }
        private void TitleBar_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                Point newPoint = e.GetPosition(this);
                Vector diff = startPoint - newPoint;

                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    DragMove();
                }
            }
        }

        // Информация об обновлении.
        private void ThisVersion_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(@"Добавлено:
- масштабирование карты.
Исправлено:
- отображение строки-меню", "Версия 3.00", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}