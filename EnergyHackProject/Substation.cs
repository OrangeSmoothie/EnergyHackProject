namespace EnergyHackProject
{
    public class SubstationClass
    {
        public string Organization { get; set; }//Организация //класс
        public string Branch { get; set; }//Филиал  //класс
        public string NumberPP { get; set; } //№ п.п.
        public string NameRES { get; set; }//Наименование РЭС
        public string NameSubstation { get; set; }//Наименование
        public string ClassVoltage { get; set; }//Класс напряжения
        public int YearInput { get; set; } //Год ввода
        //Трансформаторы
        public string TransNumber { get; set; }//№
        public string TransType { get; set; }//Тип
        public double TransNomPower { get; set; } //Номинальная мощность, МВА
        public double TransLoadWinter { get; set; } //Загрузка (зимний максимум), %
        public int TransYearManufacture { get; set; }//Год изготовления
        public int TransYearOn { get; set; } //Год включения
        public double TransPercentWear { get; set; }//% износа* 
        public string TransCondition { get; set; }//Состояние (хор., удов.,в зоне риска) 

        public string Comments { get; set; }


        public SubstationClass()
        {

        }
        public SubstationClass(string organization, string branch, string numberPP, string NameRES_,
            string nameSubstation, string classVoltage, int yearInput,
            string Transnumber, string Transtype, double TransnomPower, double TransloadWinter,
            int TransyearManufacture, int TransyearOn, double TranspercentWear, string Transcondition)
        {
            Organization = organization;
            Branch = branch;
            NumberPP = numberPP;
            NameRES = NameRES_;
            NameSubstation = nameSubstation;
            ClassVoltage = classVoltage;
            YearInput = yearInput;

            TransNumber = Transnumber;
            TransType = Transtype;
            TransNomPower = TransnomPower;
            TransLoadWinter = TransloadWinter;
            TransYearManufacture = TransyearManufacture;
            TransYearOn = TransyearOn;
            TransPercentWear = TranspercentWear;
            TransCondition = Transcondition;
        }
    }
}
