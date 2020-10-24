using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace EnergyHackProject
{
    public static class ExportToExcel
    {
        public static int globalCounter = 0;
        public static void CreatePLExcel(List<PowerLinesClass> plsList)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Add();
            excelApp.StandardFont = "Times New Roman";
            excelApp.StandardFontSize = 10;
            //создание headerа
            Excel._Worksheet sheet = (Excel.Worksheet)excelApp.ActiveSheet;
            CreateHeaderForPL(sheet);

            sheet.Activate();
            sheet.Application.ActiveWindow.SplitRow = 5;
            sheet.Application.ActiveWindow.FreezePanes = true;

            List<string> numbers = new List<string>();
            Excel.Range Cell;
            foreach (var item in plsList)
            {
                if (!numbers.Contains(item.NumberPP))
                    numbers.Add(item.NumberPP);
            }
            //Организация
            Cell = sheet.get_Range("A6", "M6").Cells;
            Cell.Merge(Type.Missing);
            Cell.Value = plsList[0].Organization;
            Cell = sheet.get_Range("A7", "M7").Cells;
            Cell.Merge(Type.Missing);
            Cell.Value = plsList[0].Branch;
            Cell = sheet.get_Range("A8", "M8").Cells;
            Cell.Merge(Type.Missing);
            Cell.Value = plsList[0].NumberPowerLines;
            int counter = 9;
            double vl50 = 0;
            foreach (string numb in numbers)
            {
                var selectedPLS = from ss in plsList
                                  where ss.NumberPP == numb
                                  select ss;
                PowerLinesClass plcTemp = null;
                foreach (var pls in selectedPLS)
                {
                    //ЛЭП

                    if (plcTemp != null && numb == plcTemp.NumberPP)
                    {
                        Cell = sheet.get_Range("A" + (counter - 1), "A" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        Cell.Value = numb;
                    }
                    else
                        sheet.Cells[counter, 2] = numb;

                    if (plcTemp != null && pls.NamePowerLines == plcTemp.NamePowerLines)
                    {

                        Cell = sheet.get_Range("B" + (counter - 1), "B" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        Cell.Value = pls.NumberPowerLines;

                        Cell = sheet.get_Range("C" + (counter - 1), "C" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        Cell.Value = pls.NamePowerLines;

                        Cell = sheet.get_Range("D" + (counter - 1), "D" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        Cell.Value = pls.Voltage;
                        //
                        Cell = sheet.get_Range("E" + (counter - 1), "E" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        Cell.Value = pls.CommissioningYear;

                        Cell = sheet.get_Range("F" + (counter - 1), "F" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        Cell.Value = pls.Cable.CountСircuit;

                        Cell = sheet.get_Range("G" + (counter - 1), "G" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        if (pls.Cable.totalLength.trackLengthALL == -1)
                            Cell.Value = "-";
                        else
                            Cell.Value = pls.Cable.totalLength.trackLengthALL;

                        Cell = sheet.get_Range("H" + (counter - 1), "H" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        if (pls.Cable.totalLength.LenghtPerСircuitALL == -1)
                            Cell.Value = "-";
                        else
                            Cell.Value = pls.Cable.totalLength.LenghtPerСircuitALL;

                        Cell = sheet.get_Range("I" + (counter - 1), "I" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        if (pls.Cable.circuitLength.trackLength == -1)
                            Cell.Value = "-";
                        else
                            Cell.Value = pls.Cable.circuitLength.trackLength;

                        Cell = sheet.get_Range("J" + (counter - 1), "J" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        if (pls.Cable.circuitLength.LenghtPerСircuit == -1)
                            Cell.Value = "-";
                        else
                            Cell.Value = pls.Cable.circuitLength.LenghtPerСircuit;
                        //
                        Cell = sheet.get_Range("L" + (counter - 1), "L" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        Cell.Value = pls.TechnicalCondition;

                        Cell = sheet.get_Range("M" + (counter - 1), "M" + counter).Cells;
                        Cell.Merge(Type.Missing);
                        Cell.Value = pls.CommissioningYear != -1 ? (DateTime.Now.Year - pls.CommissioningYear).ToString() : "-"; ;

                    }
                    else
                    {
                        sheet.Cells[counter, 2] = pls.NumberPowerLines;
                        sheet.Cells[counter, 3] = pls.NamePowerLines;
                        sheet.Cells[counter, 4] = pls.Voltage;
                        if (pls.CommissioningYear == -1)
                            sheet.Cells[counter, 5] = "-";
                        else
                            sheet.Cells[counter, 5] = pls.CommissioningYear;
                        //Кабель
                        sheet.Cells[counter, 6] = pls.Cable.CountСircuit;
                        if (pls.Cable.totalLength.trackLengthALL == -1)
                            sheet.Cells[counter, 7] = "-";
                        else
                            sheet.Cells[counter, 7] = pls.Cable.totalLength.trackLengthALL;
                        if (pls.Cable.totalLength.LenghtPerСircuitALL == -1)
                            sheet.Cells[counter, 8] = "-";
                        else
                            sheet.Cells[counter, 8] = pls.Cable.totalLength.LenghtPerСircuitALL;
                        if (pls.Cable.circuitLength.trackLength == -1)
                            sheet.Cells[counter, 9] = "-";
                        else
                            sheet.Cells[counter, 9] = pls.Cable.circuitLength.trackLength;
                        if (pls.Cable.circuitLength.LenghtPerСircuit == -1)
                            sheet.Cells[counter, 10] = "-";
                        else
                            sheet.Cells[counter, 10] = pls.Cable.circuitLength.LenghtPerСircuit;
                        //техническое состояние
                        sheet.Cells[counter, 12] = pls.TechnicalCondition;
                        //срок службы ЛЭП НА 01.01.2019
                        sheet.Cells[counter, 13] = pls.CommissioningYear != -1 ? (DateTime.Now.Year - pls.CommissioningYear).ToString() : "-";
                    }
                    //Марки кабеля
                    sheet.Cells[counter, 11] = pls.Cable.Model;
                    counter++;
                    plcTemp = pls;

                }
            }
            CreatePiesForPL(sheet);
            excelApp.Visible = true;


        }
        public static void CreateSSExcel(List<SubstationClass> subList)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Add();
            excelApp.StandardFont = "Times New Roman";
            excelApp.StandardFontSize = 10;
            //создание headerа
            Excel._Worksheet sheet = (Excel.Worksheet)excelApp.ActiveSheet;
            CreateHeaderForSS(sheet);
            sheet.Activate();
            sheet.Application.ActiveWindow.SplitRow = 2;
            sheet.Application.ActiveWindow.FreezePanes = true;
            Excel.Range Cell;
            int counter = 3;
            Excel.Range cells = sheet.get_Range("A3", "O4");
            cells.Font.Bold = false;
            SubstationClass pls;
            string prevRes = "";
            int prevResPos = -1;
            int countOfCells = 0;
            for (int i = 0; i < subList.Count; i++)
            {
                countOfCells++;
                pls = subList[i];
                if (pls.NumberPP != "-")
                    sheet.Cells[counter, 2] = pls.NumberPP;
                else
                {
                    sheet.Cells[counter, 2] = pls.Branch;
                    Cell = sheet.get_Range("A" + counter, "O" + counter).Cells;

                    Cell.Merge(Type.Missing);
                    counter++;
                    globalCounter++;
                    CreatePiesForSS(sheet, counter, countOfCells, subList);
                    continue;
                }
                if (prevRes == pls.NameRES)
                {
                    Cell = sheet.get_Range("A" + prevResPos, "A" + counter).Cells;
                    Cell.Merge(Type.Missing);
                    Cell.Value = prevRes;
                }
                else
                {
                    sheet.Cells[counter, 1] = pls.NameRES;
                    prevResPos = counter;
                    prevRes = pls.NameRES;
                }
                sheet.Cells[counter, 3] = pls.NameSubstation;
                sheet.Cells[counter, 4] = pls.ClassVoltage;
                if (pls.YearInput != -1)
                {
                    sheet.Cells[counter, 5] = pls.TransYearManufacture;
                    sheet.Cells[counter, 14] = DateTime.Now.Year - pls.TransYearManufacture; // с года ввода ПС
                }
                else
                {
                    sheet.Cells[counter, 5] = pls.TransYearManufacture;
                    sheet.Cells[counter, 14] = DateTime.Now.Year - pls.TransYearManufacture; // с года ввода ПС
                }
                //Трансформатор
                sheet.Cells[counter, 6] = pls.TransNumber;  //номер
                sheet.Cells[counter, 7] = pls.TransType;//тип
                if (pls.TransNomPower != -1)
                    sheet.Cells[counter, 8] = pls.TransNomPower;//номинальная мощность
                else
                    sheet.Cells[counter, 8] = "-";
                if (pls.TransLoadWinter != -1)
                    sheet.Cells[counter, 9] = pls.TransLoadWinter;//загрузка(зимний макс)
                else
                    sheet.Cells[counter, 9] = "-";
                if (pls.TransYearManufacture != -1)
                {
                    sheet.Cells[counter, 10] = pls.TransYearManufacture;//год изготовления
                    sheet.Cells[counter, 15] = DateTime.Now.Year - pls.TransYearManufacture; // с года ввода ПС
                }
                else
                {
                    sheet.Cells[counter, 10] = "-";//год изготовления
                    sheet.Cells[counter, 15] = "-"; // с года ввода ПС
                }
                if (pls.TransYearOn != -1)
                    sheet.Cells[counter, 11] = pls.TransYearOn;//год включения
                else
                    sheet.Cells[counter, 11] = "-";
                if (pls.TransPercentWear != -1)//% износа
                    sheet.Cells[counter, 12] = pls.TransPercentWear;//год включения
                else
                    sheet.Cells[counter, 12] = "-";
                sheet.Cells[counter, 13] = pls.TransCondition;//состояние
                counter++;
            }
            //workSheet.Columns.AutoFit();
            //workSheet.Rows.AutoFit();
            excelApp.Visible = true;

        }

        private static void CreatePiesForPL(Excel._Worksheet sheet)
        {
            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = xlCharts.Add(5, 180, 300, 250);//создаем графиксразу под записями с данными
            Excel.Chart chart1 = myChart.Chart;

            chartRange = sheet.get_Range("T2", "U5");//Задаем ячейки данных для графика
            chart1.SetSourceData(chartRange, Type.Missing);
            chart1.ChartType = Excel.XlChartType.xl3DColumn;

            chart1.ApplyCustomType(Excel.XlChartType.xl3DPie);
        }
        private static void CreatePiesForSS(Excel._Worksheet sheet, int counter, int countOfCells, List<SubstationClass> list)
        {
            Excel.Range chartRange;
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = xlCharts.Add(905, 120 + counter * 20 - countOfCells * 5, 300, 300);//создаем графиксразу под записями с данными
            Excel.Chart chart1 = myChart.Chart;

            chartRange = sheet.get_Range("Q" + (counter + 4), "R" + (counter + 7));//Задаем ячейки данных для графика
            chart1.SetSourceData(chartRange, Type.Missing);
            //chart1.ChartType = Excel.XlChartType.xl3DColumn;
            chart1.ApplyCustomType(Excel.XlChartType.xl3DPie);
            chart1.ChartTitle.Text = "Подстанции(штук, %)";
            //sheet.Cells[counter + 4, 18] = "Подстанции(штук, %)";
            Excel.Range rng = sheet.get_Range("Q" + (counter + 4), "R" + (counter + 7));
            rng.Font.Size = 11;
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            sheet.Cells[counter + 4, 17] = "подстанции";
            sheet.Cells[counter + 4, 18] = "штук";
            sheet.Cells[counter + 4, 19] = "%";
            sheet.Cells[counter + 5, 17] = "свыше 50 лет";
            sheet.Cells[counter + 6, 17] = "от 26 50 лет";
            sheet.Cells[counter + 7, 17] = "до 25 лет";
            sheet.Cells[counter + 8, 17] = "сумма";

            (sheet.Cells[counter + 5, 17] as Excel.Range).Interior.Color = Color.Red;
            (sheet.Cells[counter + 6, 17] as Excel.Range).Interior.Color = Color.LightGreen;
            (sheet.Cells[counter + 7, 17] as Excel.Range).Interior.Color = Color.LightGray;

            int first = (from li in list
                         where ((DateTime.Now.Year - li.TransYearManufacture) > 50 && li.Branch == "Сети " + globalCounter)
                         select li).Count();
            int second = (from li in list
                          where ((DateTime.Now.Year - li.TransYearManufacture) < 50 && (DateTime.Now.Year - li.TransYearManufacture) > 25 && li.Branch == "Сети " + globalCounter)
                          select li).Count();
            int third = (from li in list
                         where ((DateTime.Now.Year - li.TransYearManufacture) < 25 && li.Branch == "Сети " + globalCounter)
                         select li).Count();
            sheet.Cells[counter + 5, 18] = first;
            sheet.Cells[counter + 6, 18] = second;
            sheet.Cells[counter + 7, 18] = third;
            sheet.Cells[counter + 8, 18] = first + second + third;
            sheet.Cells[counter + 5, 19] = 100 * first / (first + second + third);
            sheet.Cells[counter + 6, 19] = 100 * second / (first + second + third);
            sheet.Cells[counter + 7, 19] = 100 * third / (first + second + third);
            sheet.Cells[counter + 8, 19] = 100;

            //трансформаторы

            sheet.Cells[counter + 4 + 7, 17] = "трансформаторы";
            sheet.Cells[counter + 4 + 7, 18] = "МВА";
            sheet.Cells[counter + 4 + 7, 19] = "%";
            sheet.Cells[counter + 5 + 7, 17] = "свыше 50 лет";
            sheet.Cells[counter + 6 + 7, 17] = "от 26 50 лет";
            sheet.Cells[counter + 7 + 7, 17] = "до 25 лет";
            sheet.Cells[counter + 8 + 7, 17] = "сумма";

            (sheet.Cells[counter + 5 + 7, 17] as Excel.Range).Interior.Color = Color.Red;
            (sheet.Cells[counter + 6 + 7, 17] as Excel.Range).Interior.Color = Color.LightGreen;
            (sheet.Cells[counter + 7 + 7, 17] as Excel.Range).Interior.Color = Color.LightGray;

            double ffirst = (from li in list
                             where ((DateTime.Now.Year - li.TransYearManufacture) > 50 && li.Branch == "Сети " + globalCounter)
                             select li.TransNomPower).Sum();
            double ssecond = (from li in list
                              where ((DateTime.Now.Year - li.TransYearManufacture) < 50 && (DateTime.Now.Year - li.TransYearManufacture) > 25 && li.Branch == "Сети " + globalCounter)
                              select li.TransNomPower).Sum();
            double tthird = (from li in list
                             where ((DateTime.Now.Year - li.TransYearManufacture) < 25 && li.Branch == "Сети " + globalCounter)
                             select li.TransNomPower).Sum();
            sheet.Cells[counter + 5 + 7, 18] = ffirst;
            sheet.Cells[counter + 6 + 7, 18] = ssecond;
            sheet.Cells[counter + 7 + 7, 18] = tthird;
            sheet.Cells[counter + 8 + 7, 18] = ffirst + ssecond + tthird;
            sheet.Cells[counter + 5 + 7, 19] = 100 * ffirst / (ffirst + ssecond + tthird);
            sheet.Cells[counter + 6 + 7, 19] = 100 * second / (ffirst + ssecond + tthird);
            sheet.Cells[counter + 7 + 7, 19] = 100 * tthird / (ffirst + ssecond + tthird);
            sheet.Cells[counter + 8 + 7, 19] = 100;

        }
        //Шаблон для заголовков ЛЭП
        private static void CreateHeaderForPL(Excel._Worksheet sheet)
        {
            Excel.Range cells = sheet.get_Range("A1", "M8");
            cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            cells.Font.Bold = true;
            cells = sheet.get_Range("A1", "E5");
            cells.Orientation = 90;
            cells = sheet.get_Range("L1", "M5");
            cells.Orientation = 90;

            //№ п.п.
            Excel.Range firstCell = sheet.get_Range("A1", "A5").Cells;
            firstCell.Merge(Type.Missing);
            firstCell.Value = "№ п.п.";
            // диспетчерский номер ЛЭП
            Excel.Range secondCell = sheet.get_Range("B1", "B5").Cells;
            secondCell.Merge(Type.Missing);
            secondCell.Value = "Диспетчерский номер ЛЭП";
            //Наименование ЛЭП
            Excel.Range thirdCell = sheet.get_Range("C1", "C5").Cells;
            thirdCell.Merge(Type.Missing);
            thirdCell.Value = "Наименование ЛЭП";
            thirdCell.Orientation = 0;
            //Напряжение, КВ
            Excel.Range fourthCell = sheet.get_Range("D1", "D5").Cells;
            fourthCell.Merge(Type.Missing);
            fourthCell.Value = "Напряжение, КВ";
            //Год ввода в эксплуатацию
            Excel.Range fifthCell = sheet.get_Range("E1", "E5").Cells;
            fifthCell.Merge(Type.Missing);
            fifthCell.Value = "Год ввода в эксплуатацию";
            //Провод, кабель
            Excel.Range sixthCell = sheet.get_Range("F1", "K1").Cells;
            sixthCell.Merge(Type.Missing);
            sixthCell.Value = "Провод, кабель";
            //Количество цепей
            Excel.Range seventhCell = sheet.get_Range("F2", "F5").Cells;
            seventhCell.Merge(Type.Missing);
            seventhCell.Orientation = 90;
            seventhCell.Value = "Количество цепей";
            //Длина всего, км
            Excel.Range Cell8 = sheet.get_Range("G2", "H2").Cells;
            Cell8.Merge(Type.Missing);
            Cell8.Value = "Длина всего, км";
            //по трассе
            Excel.Range Cell9 = sheet.get_Range("G3", "G5").Cells;
            Cell9.Merge(Type.Missing);
            Cell9.Orientation = 90;
            Cell9.Value = "по трассе";
            //на 1 цепь
            Excel.Range Cell10 = sheet.get_Range("H3", "H5").Cells;
            Cell10.Merge(Type.Missing);
            Cell10.Orientation = 90;
            Cell10.Value = "на 1 цепь";
            //Длина в т.ч. по участкам, км
            Excel.Range Cell11 = sheet.get_Range("I2", "J2").Cells;
            Cell11.Merge(Type.Missing);
            Cell11.Value = "Длина в т.ч. по участкам, км";
            //по трассе
            Excel.Range Cell12 = sheet.get_Range("I3", "I5").Cells;
            Cell12.Merge(Type.Missing);
            Cell12.Orientation = 90;
            Cell12.Value = "по трассе";
            //на 1 цепь
            Excel.Range Cell13 = sheet.get_Range("J3", "J5").Cells;
            Cell13.Merge(Type.Missing);
            Cell13.Orientation = 90;
            Cell13.Value = "на 1 цепь";
            //марка
            Excel.Range Cell14 = sheet.get_Range("K3", "K5").Cells;
            Cell14.Merge(Type.Missing);
            Cell14.Value = "марка";
            //Техническое состояние
            Excel.Range Cell15 = sheet.get_Range("L1", "L5").Cells;
            Cell15.Merge(Type.Missing);
            Cell15.Value = "Техническое состояние";
            //Срок службы ЛЭП на 01.01.2019 г.
            Excel.Range Cell16 = sheet.get_Range("M1", "M5").Cells;
            Cell16.Merge(Type.Missing);
            Cell16.Value = "Срок службы ЛЭП на 01.01.2019 г.";

            sheet.Cells[1, 16] = "КЛ";
            sheet.Cells[2, 16] = "свыше 50 лет";
            sheet.Cells[3, 16] = "от 36 до 50 лет";
            sheet.Cells[4, 16] = "до 35 лет";
            sheet.Cells[5, 16] = "всего";

            sheet.Cells[1, 17] = "км";
            sheet.Cells[1, 18] = "%";
            sheet.Cells[1, 21] = "км";
            sheet.Cells[1, 22] = "%";

            sheet.Cells[1, 20] = "ВЛ";
            sheet.Cells[2, 20] = "свыше 50 лет";
            sheet.Cells[3, 20] = "от 36 до 50 лет";
            sheet.Cells[4, 20] = "до 35 лет";
            sheet.Cells[5, 20] = "всего";

            sheet.Cells[1, 26] = "всего";
            sheet.Cells[4, 26] = "без КЛ";
            sheet.Cells[1, 27] = "км";
            sheet.Cells[1, 28] = "%";
        }
        //Шаблон для заголовком ПС
        private static void CreateHeaderForSS(Excel._Worksheet sheet)
        {
            Excel.Range cells = sheet.get_Range("A1", "O2");
            cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            cells.Font.Bold = true;
            //cells.ShrinkToFit = true;

            //Наименование РЭС
            Excel.Range Cell1 = sheet.get_Range("A1", "A2").Cells;
            Cell1.Merge(Type.Missing);
            Cell1.Value = "Наименование РЭС";
            Cell1.Orientation = 90;
            // № п.п.
            Excel.Range Cell2 = sheet.get_Range("B1", "B2").Cells;
            Cell2.Merge(Type.Missing);
            Cell2.Orientation = 90;
            Cell2.Value = "№ п.п.";
            //Подстанции
            Excel.Range Cell3 = sheet.get_Range("C1", "E1").Cells;
            Cell3.Merge(Type.Missing);
            Cell3.Value = "Подстанции";
            //Наименование
            Excel.Range Cell4 = sheet.get_Range("C2", "C2").Cells;
            Cell4.Merge(Type.Missing);
            Cell4.Value = "Наименование";
            //Класс напряжения
            Excel.Range Cell5 = sheet.get_Range("D2", "D2").Cells;
            Cell5.Merge(Type.Missing);
            Cell5.Orientation = 90;
            Cell5.Value = "Класс напряжения";
            //Год ввода
            Excel.Range Cell6 = sheet.get_Range("E2", "E2").Cells;
            Cell6.Merge(Type.Missing);
            Cell6.Orientation = 90;
            Cell6.Value = "Год ввода";
            //Трансформаторы
            Excel.Range Cell7 = sheet.get_Range("F1", "M1").Cells;
            Cell7.Merge(Type.Missing);
            Cell7.Value = "Трансформаторы";
            //№
            Excel.Range Cell8 = sheet.get_Range("F2", "F2").Cells;
            Cell8.Merge(Type.Missing);
            Cell8.Value = "№";
            //Тип
            Excel.Range Cell9 = sheet.get_Range("G2", "G2").Cells;
            Cell9.Merge(Type.Missing);
            Cell9.Value = "Тип";
            //Номинальная мощность, МВА
            Excel.Range Cell10 = sheet.get_Range("H2", "H2").Cells;
            Cell10.Merge(Type.Missing);
            Cell10.Orientation = 90;
            Cell10.Value = "Номинальная мощность, МВА";
            //Загрузка (зимний максимум), %
            Excel.Range Cell105 = sheet.get_Range("I2", "I2").Cells;
            Cell105.Merge(Type.Missing);
            Cell105.Orientation = 90;
            Cell105.Value = "Загрузка (зимний максимум), %";
            //Год изготовления
            Excel.Range Cell11 = sheet.get_Range("J2", "J2").Cells;
            Cell11.Merge(Type.Missing);
            Cell11.Orientation = 90;
            Cell11.Value = "Год изготовления";
            //Год включения
            Excel.Range Cell12 = sheet.get_Range("K2", "K2").Cells;
            Cell12.Merge(Type.Missing);
            Cell12.Orientation = 90;
            Cell12.Value = "Год включения";
            //% износа* 
            Excel.Range Cell125 = sheet.get_Range("L2", "L2").Cells;
            Cell125.Merge(Type.Missing);
            Cell125.Orientation = 90;
            Cell125.Value = "% износа*";
            //Состояние (хор., удов.,в зоне риска)
            Excel.Range Cell13 = sheet.get_Range("M2", "M2").Cells;
            Cell13.Merge(Type.Missing);
            Cell13.Orientation = 90;
            Cell13.Value = "Состояние (хор., удов.,в зоне риска)";
            //Срок службы на 01.01.2019 г. 
            Excel.Range Cell14 = sheet.get_Range("N1", "O1").Cells;
            Cell14.Merge(Type.Missing);
            Cell14.Value = "Срок службы на 01.01.2019 г. ";
            //с года ввода ПС
            Excel.Range Cell15 = sheet.get_Range("N2", "N2").Cells;
            Cell15.Merge(Type.Missing);
            Cell15.Orientation = 90;
            Cell15.Value = "с года ввода ПС";
            //с года изготов.тр - ра
            Excel.Range Cell16 = sheet.get_Range("O2", "O2").Cells;
            Cell16.Merge(Type.Missing);
            Cell16.Orientation = 90;
            Cell16.Value = "с года изготов.тр - ра";
        }
    }
}
