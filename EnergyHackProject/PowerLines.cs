namespace EnergyHackProject
{
    public class PowerLinesClass
    {
        public string Organization { get; set; }//Организация //класс
        public string Branch { get; set; }//Филиал  //класс

        public string NumberPP { get; set; } //№ п.п.

        public string NumberPowerLines { get; set; } //Диспетчерский номер ЛЭП
        public string NamePowerLines { get; set; }//Наименование ЛЭП
        public int Voltage { get; set; }//Напряжение, кВ -------------Лёха, это вроде должно быть string, так как никаких расчетов с напряжением не будет. Надо просто показать какой класс напряжения и все. 
        public int CommissioningYear { get; set; }//Год ввода в эксплуатацию
        public CableClass Cable { get; set; } //Провод, Кабель
        public string TechnicalCondition { get; set; }//Техническое состояние

        public string Comments { get; set; }

        public PowerLinesClass()
        {
            Cable = null;
        }
        public PowerLinesClass(string organization, string branch, string numberPP, string numberPowerLines,
            string namePowerLines, int voltage, int commissioningYear, CableClass cable, string technicalCondition)
        {
            Organization = organization;
            Branch = branch;
            NumberPP = numberPP;
            NumberPowerLines = numberPowerLines;
            NamePowerLines = namePowerLines;
            Voltage = voltage;
            CommissioningYear = commissioningYear;
            Cable = cable;
            TechnicalCondition = technicalCondition;
        }
    }

    public class CableClass
    {
        public int CountСircuit { get; set; } //количество цепей
        public TotalLength totalLength;
        public СircuitLength circuitLength;
        public string Model { get; set; } //марка// мб класс
        public CableClass(int countCircuit, TotalLength tl, СircuitLength cl, string model)
        {
            CountСircuit = countCircuit;
            totalLength = tl;
            circuitLength = cl;
            Model = model;

        }

        public struct TotalLength//Длина всего, км
        {
            public double trackLengthALL;//по трассе
            public double LenghtPerСircuitALL;//на 1 цепь
        }

        public struct СircuitLength//Длина в т.ч. по участкам, км
        {
            public double trackLength;//по трассе
            public double LenghtPerСircuit;//на 1 цепь
        }
    }
}
