using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.UserSkins;
using DevExpress.Skins;
using DevExpress.LookAndFeel;
using ALogic.DBConnector;
using ALogic.Logic.SPR;

namespace DXApplication1
{
    public enum eForm
    {
        Пусто = 0,
        ПарсингПрайсов = 1,
        НастройкиПарсингаПрайсов = 2,
        ЗагруженныеПрайсы = 3,
        ПрасингАссортиментаМск = 4,
        ПарсингНаличияПоставщика = 5,
        ПланПродаж = 6,
        ПланЗакупок = 7,
        ОтчетНаборкиКомплектации = 8,
        ПарсингКолвоНаПаллете = 9,
        ПарсингБухОстатков = 10,
        ОтчетКредиторскаяЗадолженность = 11,
        СправочникПриказов = 12,
        СправочникСотрудников = 13,
        ОтчетДебиторскаяЗадолженность = 14,
        РоботВыгрузкаВ1С = 15,
        ПеревыгрузкаВ1С = 16,
        Тест = 17,
        РоботПарсингаПрайсов = 18,
        РоботДиадок = 19,
        РоботОтветовПоставщика = 20,
        НастройкаДоставкиНаправления = 21,
        ОтчетоПродажах = 22,
        //БонусыАрконаНастройки = 23, заменены на новые настройки
        БонусыАрконаУправление = 24,
        ОтчетОбъемыПродаж = 25,
        БонусыАрконаСписок = 26,
        ПарсингСоСрокомДействия = 27,
        БонусыАрконаНастройкиНовые = 23,
        РоботПарсингаПрайсовA1 = 29,
        ПарсингАВСПоставщика = 30
    }
    static class Program
    {
        public static string mainConnection = @"Server=dbsrv2;Database=real;Integrated Security=SSPI;Connect Timeout=600";
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            //при сборке ботов оставляем пользователя EasZakTov и раскомментируем вызов конкретного бота в args[2]
            //при сборке десктопных приложений комментим вообще все внутри директивы #Region 
            #region Закомментировать все перед релизом

            args = new string[3];

            //args[0] = "EasZakTov";
            //args[1] = "ddfi3)es";

            args[0] = "MuhinAN";
            args[1] = "MuhinAN1017";

            //args[0] = "ZolotuhinAS";
            //args[1] = "ZolotuhinAS490";

            //args[0] = "KubrakovMA";
            //args[1] = "KubrakovMA291";

            //args[0] = "dubininaa";
            //args[1] = "dubininaa187";

            //args[0] = "gospodarikovanm";
            //args[1] = "natali123";

            //args[0] = "CheremushkinaTA";
            //args[1] = "CheremushkinaTA787";

            //args[0] = "ArhipovaEA";
            //args[1] = "ArhipovaEA143";

            // args[0] = "fedchenkoas";
            // args[1] = "kigh701";

            //args[0] = "lazarevop";
            //args[1] = "lazarevop122";

            //args[0] = "natali";
            //args[1] = "natali123";

            //args[0] = "bugakovaov";
            //args[1] = "bugakova549";

            //args[0] = "TurishevAV";
            //args[1] = "TurishevAV312";

            //args[0] = "BaranovaUS";
            //args[1] = "dontwork";

            //args[0] = "DavydenkoAV";
            //args[1] = "DavydenkoAV64";

            //args[0] = "BoykovDV";
            //args[1] = "boykovdv53";

            //args[0] = "kravtsovaa";
            // args[1] = "Zaq123";


            //args[2] = ((int)eForm.РоботПарсингаПрайсов).ToString();
            //args[2] = ((int)eForm.РоботОтветовПоставщика).ToString();
            // args[2] = ((int)eForm.РоботДиадок).ToString();       //01.08.2021 вынесен в отдельный проект, здесь его уже нет
            //args[2] = ((int)eForm.РоботВыгрузкаВ1С).ToString();
            //args[2] = ((int)eForm.РоботПарсингаПрайсовA1).ToString();

            //args[2] = ((int)eForm.ОтчетоПродажах).ToString();
            args[2] = ((int)eForm.ОтчетОбъемыПродаж).ToString();

            //args[2] = ((int)eForm.ПарсингСоСрокомДействия).ToString();
            //args[2] = ((int)eForm.ПарсингПрайсов).ToString();
            //args[2] = ((int)eForm.ПланПродаж).ToString();
            //args[2] = ((int)eForm.ПарсингАВСПоставщика).ToString();

            //args[2] = ((int)eForm.БонусыАрконаСписок).ToString();
            //args[2] = ((int)eForm.БонусыАрконаНастройки).ToString(); //c 01.01.2023 не используется, есть новые настройки
            //args[2] = ((int)eForm.БонусыАрконаУправление).ToString();
            //args[2] = ((int)eForm.СправочникСотрудников).ToString();
            //args[2] = ((int)eForm.ЗагруженныеПрайсы).ToString();
            //args[2] = ((int)eForm.ПеревыгрузкаВ1С).ToString();

            //args[2] = ((int)eForm.БонусыАрконаНастройкиНовые).ToString();


            #endregion


            //ProjectProperty.LoadDataAppConfig();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length > 0)
            {
                if (User.LoginUser(args[0], args[1]))
                {
                    //Form form = new WFMain(args[2]);
                    //WindowOpener.MainForm = (WFMain)form;
                    Application.Run(new Form1());
                }
            }
        }
    }
}
